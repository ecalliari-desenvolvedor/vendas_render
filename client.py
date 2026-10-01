# Wrapper de compatibilidade para client.py
# Redireciona para o app.py unificado para produção no Render
from app import app

if __name__ == '__main__':
    import os
    port = int(os.getenv('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=True)
