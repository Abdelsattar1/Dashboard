import { GoogleGenAI, Type } from '@google/genai';
import type Database from 'better-sqlite3';
import type { AutoFixActionType, SmartAuditReport, UserProfile } from '../src/types/bi.ts';
import {
  executeAutoFix as executeAutoFixCore,
  listAuditBackups,
  listRuntimeErrorLogs,
  recordRuntimeError,
  rollbackAutoFix as rollbackAutoFixCore,
  runSmartAudit as runSmartAuditCore,
  type RecordRuntimeErrorInput,
} from './smart-audit-core.ts';

export { listAuditBackups, listRuntimeErrorLogs, recordRuntimeError };
export type { RecordRuntimeErrorInput };

export function isGeminiKeyConfigured(): boolean {
  const raw = process.env.GEMINI_API_KEY?.trim();
  if (!raw || raw === 'MY_GEMINI_API_KEY' || raw === 'your_api_key_here') {
    return false;
  }
  return true;
}

function getGeminiClient(): GoogleGenAI | null {
  if (!isGeminiKeyConfigured()) return null;
  return new GoogleGenAI({
    apiKey: process.env.GEMINI_API_KEY!.trim(),
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

export function runSmartAudit(
  db: Database.Database,
  adminUser: UserProfile,
  options?: { auditAdminBypassNote?: boolean }
): SmartAuditReport {
  return runSmartAuditCore(db, adminUser, {
    ...options,
    aiConfigured: isGeminiKeyConfigured(),
  });
}

export function executeAutoFix(
  db: Database.Database,
  adminUser: UserProfile,
  actionType: AutoFixActionType,
  findingCode?: string
): { backupId: number; summary: string; report: SmartAuditReport } {
  return executeAutoFixCore(db, adminUser, actionType, findingCode, {
    aiConfigured: isGeminiKeyConfigured(),
  });
}

export function rollbackAutoFix(
  db: Database.Database,
  adminUser: UserProfile,
  backupId: number
): { summary: string; report: SmartAuditReport } {
  return rollbackAutoFixCore(db, adminUser, backupId, {
    aiConfigured: isGeminiKeyConfigured(),
  });
}

export async function enrichAuditWithGemini(
  db: Database.Database,
  adminUser: UserProfile
): Promise<{ usedAi: boolean; aiSummary: string; report: SmartAuditReport }> {
  const report = runSmartAudit(db, adminUser);
  const ai = getGeminiClient();

  if (!ai) {
    return {
      usedAi: false,
      aiSummary:
        'يعمل المراجع الذكي حالياً بمحرك القواعد المحلي بالكامل بدون الحاجة إلى مفتاح AI خارجي؛ جميع الفحوصات والإصلاحات التلقائية وتوليد برومبتات الإصلاح فعالة.',
      report,
    };
  }

  const conciseInput = {
    healthScore: report.healthScore,
    summaryCounts: report.summaryCounts,
    findings: report.findings.slice(0, 12).map((f) => ({
      code: f.code,
      severity: f.severity,
      title: f.title,
      description: f.description,
      affectedArea: f.affectedArea,
      details: f.details.slice(0, 4),
    })),
    runtimeErrors: report.runtimeErrors.slice(0, 8).map((e) => ({
      id: e.id,
      kind: e.kind,
      tabName: e.tabName,
      message: e.message,
      stack: e.stack.slice(0, 500),
    })),
  };

  const response = await ai.models.generateContent({
    model: 'gemini-4-hydrogen',
    contents: `حلل نتائج الفحص الذكي وأخطاء التشغيل التالية لنظام تحليلات مبيعات FMCG، وقدّم ملخصاً تنفيذياً واضحاً بالعربية مع توصيات عملية وتحليل لكل خطأ مسجل:\n${JSON.stringify(conciseInput, null, 2)}`,
    config: {
      systemInstruction:
        'أنت مهندس جودة وموثوقية أنظمة React + TypeScript + SQLite. قدّم إجابة دقيقة ومباشرة باللغة العربية بصيغة JSON مطابقة للـ schema بدون إضافة أسرار أو أوامر تخريبية.',
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          executiveSummary: {
            type: Type.STRING,
            description: 'ملخص تنفيذي واضح بالعربية لحالة النظام وأهم الأولويات (2-4 جمل).',
          },
          errorInsights: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                errorId: { type: Type.INTEGER },
                diagnosis: { type: Type.STRING },
                repairPrompt: { type: Type.STRING },
              },
            },
          },
        },
      },
    },
  });

  const rawText = (response.text ?? '{}').trim();
  const parsed = JSON.parse(rawText) as {
    executiveSummary?: string;
    errorInsights?: Array<{ errorId?: number; diagnosis?: string; repairPrompt?: string }>;
  };

  const aiSummary = parsed.executiveSummary?.trim() || report.aiSummary;

  const tx = db.transaction(() => {
    db.prepare(`UPDATE audit_runs SET ai_summary = ? WHERE id = ?`).run(aiSummary, report.runId);
    if (Array.isArray(parsed.errorInsights)) {
      for (const item of parsed.errorInsights) {
        if (typeof item.errorId === 'number' && item.diagnosis) {
          db.prepare(`UPDATE runtime_error_logs SET ai_analysis = ? WHERE id = ?`).run(
            item.diagnosis.trim(),
            item.errorId
          );
          const matchingFinding = report.findings.find((f) => f.id === `RUNTIME_ERR_${item.errorId}`);
          if (matchingFinding) {
            matchingFinding.aiInsight = item.diagnosis.trim();
            if (item.repairPrompt?.trim()) {
              matchingFinding.readyPrompt = item.repairPrompt.trim();
            }
          }
        }
      }
    }
  });
  tx();

  report.aiSummary = aiSummary;
  report.runtimeErrors = listRuntimeErrorLogs(db, false);
  return { usedAi: true, aiSummary, report };
}
