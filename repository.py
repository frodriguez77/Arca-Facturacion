"""
Capa de acceso a datos — Repository Pattern.

Actualmente usa archivos JSON como almacenamiento.
Para migrar a una base de datos (SQLite, PostgreSQL, etc.)
solo hay que reescribir este archivo; app.py no cambia nada.

Interfaz pública:
    EmpresaRepository   — CRUD de empresas
    UsuarioRepository   — CRUD de usuarios
"""

import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# Utilidad interna de JSON
# ---------------------------------------------------------------------------

def _read_json(path: str) -> list:
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _write_json(path: str, data: list) -> None:
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# EmpresaRepository
# ---------------------------------------------------------------------------

class EmpresaRepository:
    """
    Repositorio de empresas.

    Almacenamiento actual: empresas.json
    Para migrar a DB: reemplazar _load() y _save() por consultas ORM/SQL.
    """
    _path = os.path.join(BASE, 'empresas.json')

    # -- Lectura ---------------------------------------------------------------

    @classmethod
    def get_all(cls) -> list[dict]:
        """Devuelve todas las empresas."""
        return _read_json(cls._path)

    @classmethod
    def get_by_id(cls, empresa_id: str) -> dict | None:
        """Devuelve la empresa con ese ID, o None."""
        return next((e for e in cls.get_all() if e['id'] == empresa_id), None)

    @classmethod
    def get_by_cuit(cls, cuit: str) -> dict | None:
        """Devuelve la empresa con ese CUIT, o None."""
        return next((e for e in cls.get_all() if e['cuit'] == cuit), None)

    @classmethod
    def filter_by_ids(cls, ids: list[str]) -> list[dict]:
        """Devuelve solo las empresas cuyos IDs están en la lista."""
        return [e for e in cls.get_all() if e['id'] in ids]

    @classmethod
    def cuit_exists(cls, cuit: str, exclude_id: str | None = None) -> bool:
        """True si ya existe una empresa con ese CUIT (opcionalmente ignorando una)."""
        return any(
            e['cuit'] == cuit and e['id'] != exclude_id
            for e in cls.get_all()
        )

    @classmethod
    def id_exists(cls, empresa_id: str) -> bool:
        return any(e['id'] == empresa_id for e in cls.get_all())

    # -- Escritura -------------------------------------------------------------

    @classmethod
    def add(cls, empresa: dict) -> dict:
        """Agrega una empresa y la devuelve."""
        empresas = cls.get_all()
        empresas.append(empresa)
        _write_json(cls._path, empresas)
        return empresa

    @classmethod
    def update(cls, empresa_id: str, fields: dict) -> dict | None:
        """Actualiza los campos de una empresa y la devuelve. None si no existe."""
        empresas = cls.get_all()
        idx = next((i for i, e in enumerate(empresas) if e['id'] == empresa_id), None)
        if idx is None:
            return None
        empresas[idx].update(fields)
        _write_json(cls._path, empresas)
        return empresas[idx]

    @classmethod
    def delete(cls, empresa_id: str) -> bool:
        """Elimina una empresa. Devuelve True si existía."""
        empresas = cls.get_all()
        nuevas   = [e for e in empresas if e['id'] != empresa_id]
        if len(nuevas) == len(empresas):
            return False
        _write_json(cls._path, nuevas)
        return True


# ---------------------------------------------------------------------------
# UsuarioRepository
# ---------------------------------------------------------------------------

class UsuarioRepository:
    """
    Repositorio de usuarios.

    Almacenamiento actual: usuarios.json
    Para migrar a DB: reemplazar _load() y _save() por consultas ORM/SQL.

    Nota: las contraseñas se almacenan como hash (werkzeug).
    El repositorio solo mueve datos — el hashing es responsabilidad de app.py.
    """
    _path = os.path.join(BASE, 'usuarios.json')

    # -- Lectura ---------------------------------------------------------------

    @classmethod
    def get_all(cls) -> list[dict]:
        """Devuelve todos los usuarios (incluye password_hash — usar con cuidado)."""
        return _read_json(cls._path)

    @classmethod
    def get_all_public(cls) -> list[dict]:
        """Devuelve todos los usuarios sin el campo password_hash."""
        return [{k: v for k, v in u.items() if k != 'password_hash'}
                for u in cls.get_all()]

    @classmethod
    def get_by_id(cls, uid: str) -> dict | None:
        """Devuelve el usuario con ese ID, o None."""
        return next((u for u in cls.get_all() if u['id'] == uid), None)

    @classmethod
    def get_by_username(cls, username: str) -> dict | None:
        """Devuelve el usuario con ese username, o None."""
        return next((u for u in cls.get_all() if u['username'] == username), None)

    @classmethod
    def username_exists(cls, username: str, exclude_id: str | None = None) -> bool:
        """True si ya existe un usuario con ese username."""
        return any(
            u['username'] == username and u['id'] != exclude_id
            for u in cls.get_all()
        )

    @classmethod
    def id_exists(cls, uid: str) -> bool:
        return any(u['id'] == uid for u in cls.get_all())

    @classmethod
    def count(cls) -> int:
        """Cantidad total de usuarios."""
        return len(cls.get_all())

    # -- Escritura -------------------------------------------------------------

    @classmethod
    def add(cls, usuario: dict) -> dict:
        """Agrega un usuario y lo devuelve."""
        usuarios = cls.get_all()
        usuarios.append(usuario)
        _write_json(cls._path, usuarios)
        return usuario

    @classmethod
    def update(cls, uid: str, fields: dict) -> dict | None:
        """Actualiza campos de un usuario y lo devuelve. None si no existe."""
        usuarios = cls.get_all()
        idx = next((i for i, u in enumerate(usuarios) if u['id'] == uid), None)
        if idx is None:
            return None
        usuarios[idx].update(fields)
        _write_json(cls._path, usuarios)
        return usuarios[idx]

    @classmethod
    def delete(cls, uid: str) -> bool:
        """Elimina un usuario. Devuelve True si existía."""
        usuarios = cls.get_all()
        nuevos   = [u for u in usuarios if u['id'] != uid]
        if len(nuevos) == len(usuarios):
            return False
        _write_json(cls._path, nuevos)
        return True
