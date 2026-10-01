# Wrapper de compatibilidade para app_server.py
# Redireciona para o app.py unificado para produção no Render
from app import app, db, Usuario, Agendamento, Pedido, ItensPedido, DeletarUsuarios

if __name__ == '__main__':
    import os
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
