"""Lógica de negocio del Sistema de Scoring — Oficiales.

Este paquete NO depende de Flet ni de ninguna librería de interfaz: son
funciones sobre pandas que se pueden probar por separado y reutilizar desde
cualquier otra interfaz. Las vistas (paquete ``ui``) las invocan dentro de un
hilo de trabajo con ``asyncio.to_thread``.
"""
