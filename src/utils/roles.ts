import type { UserRole } from '../types/bi';

export function getRoleLabel(role: UserRole): string {
  switch (role) {
    case 'admin':
      return 'أدمن (مدير النظام)';
    case 'sector_manager':
      return 'مدير قطاع / مدير إدارة';
    case 'standard_user':
    default:
      return 'مستخدم قياسي (عرض وتصدير)';
  }
}

export function getRoleShortLabel(role: UserRole): string {
  switch (role) {
    case 'admin':
      return 'أدمن';
    case 'sector_manager':
      return 'مدير قطاع / إدارة';
    case 'standard_user':
    default:
      return 'مستخدم قياسي';
  }
}

export function getRolePermissionsHint(role: UserRole): string {
  switch (role) {
    case 'admin':
      return 'صلاحية كاملة: رفع Excel، إدارة المستخدمين، المسابقات، والأهداف';
    case 'sector_manager':
      return 'معاينة وتصدير وإدخال الأهداف ضمن القطاعات المحددة فقط';
    case 'standard_user':
    default:
      return 'معاينة وتصدير التقارير فقط ضمن القطاعات المحددة';
  }
}

export function canEditTargetsRole(role: UserRole): boolean {
  return role === 'admin' || role === 'sector_manager';
}
