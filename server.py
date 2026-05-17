#!/usr/bin/env python3
"""Servidor local para o add-in do Word. Execute antes de instalar o add-in."""

import http.server
import socketserver
import os
import sys

PORT = 3000
DIR = os.path.dirname(os.path.abspath(__file__))


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIR, **kwargs)

    def end_headers(self):
        # CORS necessário para o Word carregar os arquivos
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def log_message(self, fmt, *args):
        print(f"  {args[0]}  {args[1]}")


def main():
    manifest = os.path.join(DIR, "manifest.xml")

    print("=" * 55)
    print("  Servidor do Add-in Word")
    print("=" * 55)
    print(f"  URL:      http://localhost:{PORT}")
    print(f"  Arquivos: {DIR}")
    print()
    print("  Como instalar o add-in no Word:")
    print()
    print("  1. Com este servidor rodando, abra o Word")
    print("  2. Inserir → Suplementos → Meus Suplementos")
    print('  3. Clique em "Carregar Suplemento..."')
    print(f"  4. Selecione o arquivo:")
    print(f"     {manifest}")
    print()
    print("  Após instalar, o painel abre via:")
    print("  Inserir → Meus Suplementos → Ferramenta Word")
    print()
    print("  Requisições:")
    print("-" * 55)

    try:
        with socketserver.TCPServer(("", PORT), Handler) as httpd:
            httpd.serve_forever()
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"\n  Porta {PORT} já está em uso.")
            print("  O servidor pode já estar rodando. Tente acessar:")
            print(f"  http://localhost:{PORT}/taskpane.html")
        else:
            raise
    except KeyboardInterrupt:
        print("\n\n  Servidor encerrado.")


if __name__ == "__main__":
    main()
