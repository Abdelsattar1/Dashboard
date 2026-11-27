import React from 'react';
import { AlertTriangle, RotateCcw } from 'lucide-react';

interface ErrorBoundaryProps {
  tabLabel: string;
  onErrorCaptured?: (error: Error, componentStack: string) => void;
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  errorMessage: string;
}

export class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, errorMessage: '' };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return {
      hasError: true,
      errorMessage: error?.message || 'حدث خطأ غير متوقع أثناء عرض هذا التبويب',
    };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    if (this.props.onErrorCaptured) {
      this.props.onErrorCaptured(error, errorInfo?.componentStack || '');
    }
  }

  componentDidUpdate(prevProps: ErrorBoundaryProps): void {
    if (prevProps.tabLabel !== this.props.tabLabel && this.state.hasError) {
      this.setState({ hasError: false, errorMessage: '' });
    }
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="bg-white border border-rose-200 rounded-xl p-8 max-w-xl mx-auto my-8 text-center space-y-4 shadow-sm">
          <div className="w-12 h-12 rounded-full bg-rose-50 border border-rose-200 text-rose-600 flex items-center justify-center mx-auto">
            <AlertTriangle className="w-6 h-6" />
          </div>
          <div className="space-y-1.5">
            <h2 className="text-base font-bold text-slate-900">
              تعذر عرض محتوى «{this.props.tabLabel}» مؤقتاً
            </h2>
            <p className="text-xs text-slate-600 leading-relaxed">
              تم تسجيل تفاصيل الخطأ تلقائياً في سجل المراقبة لمراجعة مدير النظام. يمكنك المحاولة مرة أخرى الآن.
            </p>
            <div className="mt-2 p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-[11px] font-mono text-rose-700 text-left break-words" dir="ltr">
              {this.state.errorMessage}
            </div>
          </div>
          <button
            type="button"
            onClick={() => this.setState({ hasError: false, errorMessage: '' })}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>إعادة المحاولة</span>
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
