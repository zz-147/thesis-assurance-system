"""tas Web 双入口（本地，无账号）。"""

from .app import create_server, serve

__all__ = ["create_server", "serve"]
