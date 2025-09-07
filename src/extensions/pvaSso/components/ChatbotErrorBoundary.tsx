import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import styles from '../styles/PvaSsoApplicationCustomizer.module.scss';

interface ChatbotErrorBoundaryState {
  hasError: boolean;
  error?: Error;
  errorInfo?: React.ErrorInfo;
}

interface ChatbotErrorBoundaryProps {
  children: React.ReactNode;
  fallbackMessage?: string;
}

export class ChatbotErrorBoundary extends React.Component<
  ChatbotErrorBoundaryProps,
  ChatbotErrorBoundaryState
> {
  constructor(props: ChatbotErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(error: Error): ChatbotErrorBoundaryState {
    return {
      hasError: true,
      error
    };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    this.setState({
      error,
      errorInfo
    });

    if (process.env.NODE_ENV !== 'production') {
      console.error('Chatbot Error Boundary caught an error:', error, errorInfo);
    }
  }

  private handleRetry = (): void => {
    this.setState({ hasError: false, error: undefined, errorInfo: undefined });
  };

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div className={styles.errorBoundaryContainer}>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={true}
            className={styles.msMessageBar}
          >
            <strong>
              {this.props.fallbackMessage || 'Sorry, the chatbot encountered an error.'}
            </strong>
            <br />
            {process.env.NODE_ENV !== 'production' && this.state.error && (
              <details className={styles.errorDetails}>
                <summary>Error Details (Development Only)</summary>
                <pre className={styles.errorDetailsContent}>
                  {this.state.error.toString()}
                  {this.state.errorInfo?.componentStack}
                </pre>
              </details>
            )}
          </MessageBar>
          <div className={styles.errorRetryContainer}>
            <PrimaryButton
              onClick={this.handleRetry}
              text="Try Again"
            />
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

export default ChatbotErrorBoundary;