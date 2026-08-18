import { Component, ReactNode } from 'react';

// ErrorBoundary para capturar erros de renderização e exibir na tela
export class ErrorBoundary extends Component<{children: ReactNode}, {error: Error | null}> {
  constructor(props: {children: ReactNode}) {
    super(props);
    this.state = {error: null};
  }
  static getDerivedStateFromError(error: Error) {
    return {error};
  }
  componentDidCatch(error: Error, info: any) {
    console.error('🔴 ErrorBoundary capturou erro:', error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{
          fontFamily: 'monospace',
          padding: '32px',
          background: '#fee2e2',
          minHeight: '100vh',
          color: '#991b1b'
        }}>
          <h1 style={{fontSize: '24px', fontWeight: 'bold', marginBottom: '16px'}}>
            ❌ Erro ao renderizar o aplicativo
          </h1>
          <p style={{marginBottom: '12px', fontWeight: '600'}}>
            {this.state.error.message}
          </p>
          <pre style={{
            background: '#fca5a5',
            padding: '16px',
            borderRadius: '8px',
            overflow: 'auto',
            fontSize: '13px',
            whiteSpace: 'pre-wrap'
          }}>
            {this.state.error.stack}
          </pre>
          <button
            onClick={() => this.setState({error: null})}
            style={{
              marginTop: '16px',
              padding: '8px 20px',
              background: '#dc2626',
              color: 'white',
              border: 'none',
              borderRadius: '8px',
              cursor: 'pointer',
              fontWeight: 'bold'
            }}
          >
            Tentar novamente
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
