¡Excelente idea! Por seguridad, esa herramienta no debería estar a la vista de todos.

Para lograr esto, tenemos que hacer un pequeño cambio lógico en tu código. Actualmente, el generador de contraseñas está **arriba** del inicio de sesión, por lo que la página se lo muestra a todo el mundo antes de saber quiénes son.

Para que solo lo vean `jnavarrete` y `admin`, debemos **mover ese bloque de código hacia abajo** (después de que el usuario ingresa su clave) y encerrarlo en una condición `if` (si).

### Cómo hacerlo paso a paso:

**1. Borra el generador de la parte de arriba:**
Busca este bloque de código (que está justo debajo de `st.set_page_config(...)`) y **BÓRRALO** o córtalo:

```python
# =========================================================
# 🛠️ HERRAMIENTA TEMPORAL PARA ENCRIPTAR CONTRASEÑAS
# =========================================================
with st.expander("🛠️ Admin: Generador de Contraseñas Seguras (Desplegar)"):
    # ... todo lo que hay dentro ...

```

**2. Pégalo más abajo, protegido con el condicional:**
Ahora baja en tu código hasta pasar la zona donde el inicio de sesión fue exitoso. Pega este nuevo bloque justo debajo del botón de "Cerrar Sesión" de la barra lateral:

```python
# =========================================================
# (A PARTIR DE AQUÍ, SOLO SE EJECUTA SI EL LOGIN FUE EXITOSO)
# =========================================================

# Botón de cerrar sesión en la barra lateral
authenticator.logout("Cerrar Sesión", "sidebar")
st.sidebar.markdown(f"👋 Hola, **{st.session_state['name']}**")
st.sidebar.markdown("---")

# =========================================================
# 🛠️ HERRAMIENTA PROTEGIDA: GENERADOR DE CONTRASEÑAS
# =========================================================
# st.session_state["username"] lee el nombre de usuario de quien inició sesión
if st.session_state["username"] in ["admin", "jnavarrete"]:
    with st.expander("🛠️ Admin: Generador de Contraseñas Seguras"):
        st.info("Escribe la contraseña que quieres asignarle a un usuario. La herramienta te dará el código encriptado (Hash).")
        clave_nueva = st.text_input("Contraseña normal (Ej: chile2026):")
        
        if clave_nueva:
            import bcrypt
            salt = bcrypt.gensalt()
            hash_generado = bcrypt.hashpw(clave_nueva.encode('utf-8'), salt).decode('utf-8')
            
            st.code(hash_generado)
            st.warning("☝️ Copia el código de arriba y pégalo en el bloque de 'credentials' en tu código.")

# ... Y AQUÍ CONTINÚA EL RESTO DE TU CÓDIGO (URLs, Conexión a Sheets, Menú, etc.) ...

```

### ¿Qué hace este código?

La línea clave es `if st.session_state["username"] in ["admin", "jnavarrete"]:`.
Esto le dice a la página: *"Revisa quién acaba de iniciar sesión. Si su nombre de usuario es 'admin' o es 'jnavarrete', dibuja el expansor en la pantalla. Si es cualquier otro usuario (como 'jhernandez' o 'pperez'), sáltate este paso y no le muestres nada"*.

¡Pruébalo! Entra con Jorge y verás que no aparece. Luego cierra sesión, entra con el Admin o Juan y verás la herramienta mágicamente de vuelta.
