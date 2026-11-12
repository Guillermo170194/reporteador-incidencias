"""Edición puntual de incidencias existentes, sin inserciones ni cambios masivos."""
import streamlit as st

# Columnas capturadas en la tabla incidencias; los identificadores y metadatos no se editan.
CAMPOS = {
    "orden_buscada": "Orden buscada", "orden_suministro": "Orden de suministro",
    "tipo_entrega": "Tipo de entrega", "entidad": "Entidad",
    "almacen_clues_destino": "Almacén / CLUES destino", "clues_destino": "CLUES destino",
    "unidad_destino": "Unidad destino", "proveedor": "Proveedor",
    "orden": "Orden", "clave_cnis": "Clave CNIS", "descripcion": "Descripción",
    "piezas_emitidas": "Piezas emitidas", "piezas_recibidas_ol": "Piezas recibidas OL",
    "piezas_entregadas_clues": "Piezas entregadas CLUES", "tipo_red": "Tipo de red",
    "grupo_terapeutico": "Grupo terapéutico", "estatus_operativo": "Estatus operativo",
    "estatus_base": "Estatus base", "origen_compendio": "Origen compendio",
    "operador_logistico": "Operador logístico", "estatus_recepcion_ol": "Estatus recepción OL",
    "estatus_entrega_estado": "Estatus entrega estado",
    "estatus_incidencia_completa": "Estatus incidencia completa",
    "estatus_seguimiento": "Estatus seguimiento", "tipo_incidencia": "Tipo incidencia",
    "atribuible_a": "Atribuible a", "estatus_incidencia": "Estatus incidencia",
    "responsable": "Responsable", "observaciones": "Observaciones",
    "pdf_cedula_rechazo": "PDF cédula rechazo (enlace)",
    "pdf_correo_seguimiento": "PDF correo seguimiento (enlace)",
}


def _texto(valor):
    return "" if valor is None else str(valor)


def _leer_orden(cliente, orden):
    # Dos consultas exactas, sin interpolar filtros PostgREST ni recorrer toda la base.
    encontrados = {}
    for columna in ("orden_suministro", "orden"):
        respuesta = (cliente.table("incidencias").select("*")
                     .eq(columna, orden).limit(100).execute())
        for fila in (respuesta.data or []):
            if fila.get("id") is not None:
                encontrados[str(fila["id"])] = fila
    return list(encontrados.values())


def mostrar_editor(cliente):
    st.subheader("Editar incidencia")
    st.caption("Busca exclusivamente por orden de suministro. Solo se modifica el registro seleccionado.")
    with st.form("buscar_incidencia_por_orden"):
        orden = st.text_input("Orden de suministro", placeholder="IMB-30-02-2026-30223066-U013").strip()
        buscar = st.form_submit_button("Buscar orden", type="primary")
    if buscar:
        st.session_state["editor_orden"] = orden
        st.session_state.pop("editor_filas", None)
        if orden:
            try:
                st.session_state["editor_filas"] = _leer_orden(cliente, orden)
            except Exception as exc:
                st.error(f"No fue posible consultar las incidencias: {exc}")
    filas = st.session_state.get("editor_filas")
    if filas is None:
        return
    if not filas:
        st.warning("No hay incidencias para esa orden exacta.")
        return
    st.success(f"Se encontraron {len(filas)} incidencia(s).")
    opciones = {str(f["id"]): f for f in filas}
    seleccionado = st.selectbox(
        "Selecciona la incidencia exacta",
        list(opciones),
        format_func=lambda id_: (
            f"ID {id_} | Clave { _texto(opciones[id_].get('clave_cnis')) } | "
            f"{_texto(opciones[id_].get('tipo_incidencia'))} | "
            f"{_texto(opciones[id_].get('estatus_incidencia'))}"
        ),
    )
    original = opciones[seleccionado]
    st.info("Se conserva el ID, la fecha original y las evidencias existentes salvo que edites expresamente sus enlaces.")
    with st.form(f"editar_incidencia_{seleccionado}"):
        nuevos = {}
        for campo, etiqueta in CAMPOS.items():
            if campo not in original:
                continue
            valor = _texto(original.get(campo))
            if campo == "atribuible_a":
                alternativas = ["Estado", "Proveedor", "Operador logístico", "Otro"]
                opciones_atrib = [valor] + [x for x in alternativas if x.casefold() != valor.casefold()] if valor else alternativas
                nuevos[campo] = st.selectbox(etiqueta, opciones_atrib, key=f"ed_{seleccionado}_{campo}")
            elif campo in ("observaciones", "descripcion"):
                nuevos[campo] = st.text_area(etiqueta, value=valor, key=f"ed_{seleccionado}_{campo}")
            else:
                nuevos[campo] = st.text_input(etiqueta, value=valor, key=f"ed_{seleccionado}_{campo}")
        confirmar = st.checkbox("Confirmo que deseo actualizar esta incidencia existente")
        guardar = st.form_submit_button("Guardar cambios", type="primary")
    if not guardar:
        return
    cambios = {k: (v if v.strip() else None) for k, v in nuevos.items()
               if _texto(original.get(k)) != v}
    if not cambios:
        st.info("No hay cambios para guardar.")
        return
    if not confirmar:
        st.warning("Marca la casilla de confirmación antes de guardar.")
        return
    try:
        # Evitar sobrescribir cambios concurrentes hechos por otro usuario.
        actual = (cliente.table("incidencias").select("*")
                  .eq("id", original["id"]).limit(1).execute()).data
        if not actual:
            st.error("La incidencia ya no existe. No se guardó ningún cambio.")
            return
        if any(_texto(actual[0].get(k)) != _texto(original.get(k)) for k in cambios):
            st.error("Otro usuario modificó uno de estos campos. Busca la orden de nuevo antes de guardar.")
            return
        respuesta = (cliente.table("incidencias").update(cambios)
                     .eq("id", original["id"]).execute())
        if not respuesta.data:
            st.error("Supabase no confirmó la actualización. Revisa los permisos de la tabla.")
            return
        st.cache_data.clear()
        st.session_state["editor_filas"] = _leer_orden(cliente, st.session_state["editor_orden"])
        st.success(f"Incidencia {original['id']} actualizada. Campos modificados: {', '.join(CAMPOS[k] for k in cambios)}")
    except Exception as exc:
        st.error(f"No se pudo guardar: {exc}")
