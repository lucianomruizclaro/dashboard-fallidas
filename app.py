import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

# ==========================================
# CONFIGURACIÓN INICIAL
# ==========================================
# Forzamos a que el menú lateral inicie siempre abierto
st.set_page_config(page_title="Dashboard de Fallidas", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
/* 1. TIPOGRAFÍA PREMIUM (GOOGLE FONTS) */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');

html, body, [class*="css"], [class*="st-"], p, h1, h2, h3 {
    font-family: 'Inter', sans-serif !important;
}

/* ========================================== */
/* MARCA BLANCA Y FONDO DE LA APP             */
/* ========================================== */

/* Ocultar header, menú superior y footer de Streamlit */
header {display: none !important;}
#MainMenu {display: none !important;}
footer {display: none !important;}

/* Darle un fondo gris perla a la aplicación para que resalten las tarjetas */
.stApp {
    background-color: #F4F7F6 !important;
}

/* Asegurar que el fondo del menú lateral sea completamente blanco para contrastar */
[data-testid="stSidebar"] {
    background-color: #ffffff !important;
    border-right: 1px solid #e0e0e0 !important;
}

/* ========================================== */
/* MENÚ LATERAL (RECTÁNGULOS DEL MISMO TAMAÑO)*/
/* ========================================== */

div[data-testid="stRadio"],
div[data-testid="stRadio"] > div,
div[data-testid="stRadio"] div[role="radiogroup"] {
    width: 100% !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label > div:first-child {
    display: none !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label {
    background-color: #f8f9fa !important;
    border: 2px solid transparent !important;
    padding: 10px 15px !important;
    border-radius: 6px !important;
    margin-bottom: 5px !important;
    cursor: pointer !important;
    width: 100% !important;
    box-sizing: border-box !important;
    transition: background-color 0.2s ease, border-color 0.2s ease !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label:hover {
    background-color: #ffe8e8 !important;
    border-color: #FF4B4B !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) {
    background-color: #FF4B4B !important;
    border-color: #FF4B4B !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label p {
    color: #444 !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    margin: 0 !important;
}

div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) p {
    color: white !important;
    font-weight: 900 !important;
}

/* ========================================== */
/* MENÚ FIJO (OCULTAR BOTONES DE COLAPSAR)    */
/* ========================================== */

/* Eliminamos por completo la opción de cerrar/abrir el menú */
[data-testid="stSidebarCollapseButton"],
[data-testid="stCollapsedControl"] {
    display: none !important;
}

/* ========================================== */
/* ANIMACIONES DE ENTRADA (CSS PURO, 0 LAG)   */
/* ========================================== */

/* Animación para los gráficos de Plotly */
[data-testid="stPlotlyChart"] {
    transform-origin: bottom center !important;
    animation: chartGrow 0.7s cubic-bezier(0.175, 0.885, 0.32, 1.1) 0.3s both !important;
}

@keyframes chartGrow {
    0% { transform: translateY(30px) scaleY(0.9); opacity: 0; }
    100% { transform: translateY(0) scaleY(1); opacity: 1; }
}

@keyframes slideUpFade {
    0% { transform: translateY(30px); opacity: 0; }
    100% { transform: translateY(0); opacity: 1; }
}

@keyframes numberPop {
    0% { transform: scale(0.5) translateY(15px); opacity: 0; color: #ffcccc; }
    80% { transform: scale(1.05) translateY(-2px); opacity: 1; }
    100% { transform: scale(1) translateY(0); opacity: 1; }
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIÓN: TÍTULO DE BLOQUE (CABECERA)
# ==========================================
def mostrar_titulo_bloque(titulo):
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #FF4B4B, #e63946); padding: 18px 25px; border-radius: 12px; text-align: center; margin-bottom: 30px; box-shadow: 0px 6px 15px rgba(255, 75, 75, 0.25); animation: slideUpFade 0.5s ease-out both;">
        <h2 style="color: white; margin: 0; font-size: 32px; font-weight: 900; letter-spacing: 1px; text-transform: uppercase;">{titulo}</h2>
    </div>
    """, unsafe_allow_html=True)

st.title("Comparativa de Fallidas 2026") 

# ==========================================
# CACHÉ PARA LECTURA Y LIMPIEZA DE DATOS
# ==========================================
@st.cache_data
def cargar_y_limpiar_datos(ruta):
    xls = pd.ExcelFile(ruta, engine='openpyxl')
    nombres_hojas = xls.sheet_names 
    
    nombres_meses = {
        1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
        5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
        9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
    }
    
    mapa_meses = {}
    meses_ordenados = [] 
    
    for hoja in nombres_hojas:
        match = re.search(r'\d+', hoja) 
        if match and 1 <= int(match.group()) <= 12:
            nombre_traducido = nombres_meses[int(match.group())]
            mapa_meses[hoja] = nombre_traducido
            if nombre_traducido not in meses_ordenados:
                meses_ordenados.append(nombre_traducido)
        else:
            mapa_meses[hoja] = hoja
            if hoja not in meses_ordenados:
                meses_ordenados.append(hoja)

    df_08, df_09, df_cancel = None, None, None
    col_desvio = 'DESVIO'
    col_region = 'REGION'
    col_agente = 'GROUP_COMER' 

    if 'Base 08' in nombres_hojas and 'Base 09' in nombres_hojas:
        df_08 = pd.read_excel(xls, sheet_name='Base 08')
        df_09 = pd.read_excel(xls, sheet_name='Base 09')
        
        for df in [df_08, df_09]:
            if col_desvio in df.columns:
                df[col_desvio] = df[col_desvio].astype(str).str.strip().str.upper()
            if col_region in df.columns:
                df[col_region] = df[col_region].astype(str).str.strip().str.upper()
            if col_agente in df.columns:
                df[col_agente] = df[col_agente].astype(str).str.strip().str.upper()

    if 'Cancel' in nombres_hojas:
        df_cancel = pd.read_excel(xls, sheet_name='Cancel')
        col_canal = 'CANAL'
        col_motivo = 'MOTIVO CANCELACION'
        col_obs = 'OBS'
        
        if col_canal in df_cancel.columns:
            df_cancel[col_canal] = df_cancel[col_canal].astype(str).str.strip().str.upper()
        if col_motivo in df_cancel.columns:
            df_cancel[col_motivo] = df_cancel[col_motivo].astype(str).str.strip().str.upper()
        if col_obs in df_cancel.columns:
            df_cancel[col_obs] = df_cancel[col_obs].astype(str).str.strip().str.upper()

    return nombres_hojas, mapa_meses, meses_ordenados, df_08, df_09, df_cancel

# ==========================================
# LÓGICA PRINCIPAL Y MENÚ LATERAL
# ==========================================
ruta_excel = 'Fallidas 082026 Arg.xlsx'

try:
    nombres_hojas, mapa_meses, meses_ordenados, df_08, df_09, df_cancel = cargar_y_limpiar_datos(ruta_excel)
    
    col_desvio, col_region, col_agente = 'DESVIO', 'REGION', 'GROUP_COMER'
    col_canal, col_motivo, col_obs = 'CANAL', 'MOTIVO CANCELACION', 'OBS'

    with st.sidebar:
        st.markdown("<h2 style='text-align: center; color: #333; margin-bottom: 20px;'>Menú</h2>", unsafe_allow_html=True)
        
        seleccion = st.radio(
            "Navegación",
            [
                "1. KPIs Generales", 
                "2. Comparativa Desvíos", 
                "3. Regiones Críticas", 
                "4. Top Canales", 
                "5. Evolución Mensual",
                "6. Canceladas por Canal",
                "7. Motivos Cancelación",
                "8. Desiste del Servicio",
                "9. Autogestión"
            ],
            label_visibility="collapsed" 
        )

    if df_08 is not None and df_09 is not None:
        
        # ==========================================
        # DIAPOSITIVA 1: KPIs COMPARATIVOS 
        # ==========================================
        if seleccion == "1. KPIs Generales":
            mostrar_titulo_bloque("KPIs Generales y Comparativa")
            
            def contar_desvios(df, desvio_nombre):
                return (df[col_desvio] == desvio_nombre).sum()

            total_09 = len(df_09)
            total_08 = len(df_08)

            motivos = ['CLIENTE', 'CLIENTE AUSENTE', 'COMERCIAL', 'CONTRATISTA', 'REDES', 'SISTEMAS', 'ZONA PELIGROSA']
            counts_08 = {m: contar_desvios(df_08, m) for m in motivos}
            counts_09 = {m: contar_desvios(df_09, m) for m in motivos}

            nombres_display = {
                'CLIENTE': 'Cliente', 'CLIENTE AUSENTE': 'Ausente', 'COMERCIAL': 'Comercial', 
                'CONTRATISTA': 'Contratista', 'REDES': 'Redes', 'SISTEMAS': 'Sistemas', 'ZONA PELIGROSA': 'Z. Peligrosa'
            }

            motivos_ordenados = sorted(motivos, key=lambda m: counts_09[m], reverse=True)

            def crear_tarjeta_kpi(titulo, valor_actual, valor_anterior=None, es_total=False, delay=0.0):
                if es_total:
                    color_borde_top, color_fondo, color_borde, color_numero = "#343a40", "#ffffff", "#e0e0e0", "#111111"
                else:
                    color_borde_top, color_fondo, color_borde, color_numero = "#FF4B4B", "#ffffff", "#f5c6c6", "#FF4B4B"

                if valor_anterior is not None:
                    delta_pct = 0 if valor_anterior == 0 else ((valor_actual - valor_anterior) / valor_anterior) * 100
                    if delta_pct > 0: color_delta, flecha = "#FF4B4B", "🔺"
                    elif delta_pct < 0: color_delta, flecha = "#09AB3B", "🔻"
                    else: color_delta, flecha = "#888888", "➖"
                    
                    html_delta = f'<span style="color: {color_delta}; font-size: 22px; font-weight: bold;">{flecha} {abs(delta_pct):.1f}%</span> <span style="font-size: 18px; color: #666;">vs Agosto</span>'
                    bloque_vs = f'<div style="background-color: #f8f9fa; border-radius: 8px; padding: 10px; border: 1px solid #eee; margin-top: 15px;">{html_delta}</div>'
                else:
                    bloque_vs = f'<div style="visibility: hidden; background-color: transparent; border-radius: 8px; padding: 10px; border: 1px solid transparent; margin-top: 15px;"><span style="font-size: 22px; font-weight: bold;">➖ 0.0%</span> <span style="font-size: 18px;">vs Agosto</span></div>'

                return f"""
                <div style="background-color: {color_fondo}; padding: 30px; border-radius: 16px; text-align: center; border: 1px solid {color_borde}; border-top: 12px solid {color_borde_top}; box-shadow: 4px 4px 15px rgba(0,0,0,0.05); height: 100%; animation: slideUpFade 0.6s ease-out {delay}s both;">
                    <p style="margin: 0; font-size: 20px; color: #555; font-weight: bold; text-transform: uppercase;">{titulo}</p>
                    <p style="margin: 10px 0 0 0; font-size: 60px; color: {color_numero}; font-weight: 900; animation: numberPop 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.2) {delay + 0.15}s both;">{valor_actual}</p>
                    {bloque_vs}
                </div>
                """

            row1 = st.columns(4)
            with row1[0]: st.markdown(crear_tarjeta_kpi("Total Agosto", total_08, es_total=True, delay=0.1), unsafe_allow_html=True)
            with row1[1]: st.markdown(crear_tarjeta_kpi("Total Sept", total_09, total_08, es_total=True, delay=0.2), unsafe_allow_html=True)
            m1, m2 = motivos_ordenados[0], motivos_ordenados[1]
            with row1[2]: st.markdown(crear_tarjeta_kpi(nombres_display[m1], counts_09[m1], counts_08[m1], delay=0.3), unsafe_allow_html=True)
            with row1[3]: st.markdown(crear_tarjeta_kpi(nombres_display[m2], counts_09[m2], counts_08[m2], delay=0.4), unsafe_allow_html=True)

            st.write("<br>", unsafe_allow_html=True)

            row2 = st.columns(5)
            for i in range(5):
                m = motivos_ordenados[i + 2] 
                with row2[i]: 
                    st.markdown(crear_tarjeta_kpi(nombres_display[m], counts_09[m], counts_08[m], delay=0.5 + (i*0.1)), unsafe_allow_html=True)

        # ==========================================
        # DIAPOSITIVA 2: GRÁFICO COMPARATIVO
        # ==========================================
        elif seleccion == "2. Comparativa Desvíos":
            mostrar_titulo_bloque("Comparativa de Desvíos")
            
            df_08_agrupado = df_08[col_desvio].value_counts().reset_index()
            df_08_agrupado.columns = ['DESVIO', 'CANTIDAD']
            df_08_agrupado['MES'] = 'Base 08 (Agosto)'
            
            df_09_agrupado = df_09[col_desvio].value_counts().reset_index()
            df_09_agrupado.columns = ['DESVIO', 'CANTIDAD']
            df_09_agrupado['MES'] = 'Base 09 (Septiembre)'
            
            df_comparativo = pd.concat([df_08_agrupado, df_09_agrupado])
            df_comparativo['TEXTO_BARRA'] = df_comparativo['CANTIDAD'].apply(lambda x: f"<b>{x}</b>")
            
            fig_comp = px.bar(
                df_comparativo, x='DESVIO', y='CANTIDAD', color='MES', barmode='group',
                text='TEXTO_BARRA', color_discrete_sequence=['#adb5bd', '#FF4B4B'] 
            )
            fig_comp.update_traces(textposition='outside', textfont=dict(size=22, color='black'))
            fig_comp.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=700,
                xaxis=dict(showgrid=False, title="", tickfont=dict(size=18, color='black', family="Inter")),
                yaxis=dict(gridcolor='#E0E0E0', title="<b>Volumen de Casos</b>", title_font=dict(size=20), tickfont=dict(size=16, color='black')),
                legend=dict(font=dict(size=20, color='black'), title_text='', orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
            )
            
            st.markdown('<div style="background-color: white; padding: 20px; border-radius: 16px; box-shadow: 0px 4px 15px rgba(0,0,0,0.05); border: 1px solid #eee; animation: slideUpFade 0.7s ease-out 0.2s both;">', unsafe_allow_html=True)
            st.plotly_chart(fig_comp, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        # ==========================================
        # DIAPOSITIVA 3: TOP 3 REGIONES CON MAYOR % COMERCIAL 
        # ==========================================
        elif seleccion == "3. Regiones Críticas":
            mostrar_titulo_bloque("Top 3 Regiones Críticas")
            
            if col_region in df_09.columns and col_desvio in df_09.columns:
                totales_por_region = df_09[col_region].value_counts()
                df_comercial = df_09[df_09[col_desvio] == 'COMERCIAL']
                comerciales_por_region = df_comercial[col_region].value_counts()
                
                pct_comercial_region = (comerciales_por_region / totales_por_region * 100).fillna(0)
                top_3_regiones = pct_comercial_region.nlargest(3)
                
                cols_top = st.columns(3)
                colores_top = ['#FF4B4B', '#fd7e14', '#ffc107'] 
                
                for i, (region, pct_comercial) in enumerate(top_3_regiones.items()):
                    delay = (i + 1) * 0.2
                    cantidad_total = totales_por_region[region]
                    df_region = df_09[df_09[col_region] == region]
                    df_motivos_region = df_region[col_desvio].value_counts().reset_index()
                    df_motivos_region.columns = ['MOTIVO', 'CANTIDAD']
                    df_motivos_region['PORCENTAJE'] = (df_motivos_region['CANTIDAD'] / cantidad_total) * 100
                    df_motivos_region = df_motivos_region.nlargest(3, 'PORCENTAJE').copy()
                    df_motivos_region = df_motivos_region.sort_values('PORCENTAJE', ascending=True)
                    
                    df_motivos_region['TEXTO_BARRA'] = df_motivos_region.apply(lambda row: f"<b>{row['PORCENTAJE']:.1f}%</b>", axis=1)
                    df_motivos_region['MOTIVO_BOLD'] = df_motivos_region['MOTIVO'].apply(lambda x: f"<b>{x}</b>")
                    
                    colores_barras = ['#d3d3d3'] * len(df_motivos_region)
                    if len(colores_barras) > 0: colores_barras[-1] = colores_top[i] 
                    
                    html_region = f"""
                    <div style="border-top: 12px solid {colores_top[i]}; padding: 30px 20px; background-color: #ffffff; border-radius: 0 0 12px 12px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); animation: slideUpFade 0.6s ease-out {delay}s both;">
                        <h2 style="margin-bottom: 10px; color: #111; font-size: 38px; font-weight: 900; text-align: center;">#{i+1} {region}</h2>
                        <p style="font-size: 20px; color: #666; margin-top: 5px; margin-bottom: 0px; text-align: center;">
                            Desvío Comercial: <strong style="color: {colores_top[i]}; font-size: 34px; animation: numberPop 0.5s ease-out {delay + 0.2}s both; display: inline-block;">{pct_comercial:.1f}%</strong>
                        </p>
                        <p style="font-size: 18px; color: #888; margin-top: 5px; text-align: center;">
                            Total fallidas en región: <strong>{cantidad_total}</strong>
                        </p>
                    </div>
                    """
                    with cols_top[i]:
                        st.markdown(html_region, unsafe_allow_html=True)
                        fig_mini = go.Figure(go.Bar(
                            x=df_motivos_region['PORCENTAJE'], y=df_motivos_region['MOTIVO_BOLD'], orientation='h',
                            text=df_motivos_region['TEXTO_BARRA'], textposition='outside',
                            textfont=dict(size=20, color='black', family="Inter"),
                            marker_color=colores_barras, cliponaxis=False, hovertemplate="<b>%{y}</b>: %{x:.1f}%<extra></extra>"
                        ))
                        max_x = df_motivos_region['PORCENTAJE'].max() if not df_motivos_region.empty else 100
                        fig_mini.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=400,
                            xaxis=dict(showgrid=False, showticklabels=False, title="", range=[0, max_x * 1.35]),
                            yaxis=dict(showgrid=False, title="", tickfont=dict(size=16, color='black')), 
                            margin=dict(l=0, r=10, t=30, b=0),
                            hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
                        )
                        st.plotly_chart(fig_mini, use_container_width=True, config={'displayModeBar': False})
            else:
                st.warning("No se encontraron las columnas necesarias para el Top 3 de Regiones.")

        # ==========================================
        # DIAPOSITIVA 4: TOP CANALES
        # ==========================================
        elif seleccion == "4. Top Canales":
            mostrar_titulo_bloque("Top Canales")
            
            if col_agente in df_09.columns and col_agente in df_08.columns:
                agentes_09 = df_09[col_agente].value_counts()
                agentes_08 = df_08[col_agente].value_counts()
                top_3_agentes = agentes_09.nlargest(3)
                cols_agentes = st.columns(3)
                colores_alerta = ['#FF4B4B', '#fd7e14', '#ffc107'] 
                
                for i, (agente, cantidad_09) in enumerate(top_3_agentes.items()):
                    delay = (i + 1) * 0.2
                    cantidad_08 = agentes_08.get(agente, 0)
                    if cantidad_08 == 0:
                        html_delta = '<span style="color: #FF4B4B; font-size: 22px; font-weight: bold;">🔺 N/A (No tuvo fallidas en Agosto)</span>'
                    else:
                        delta_pct = ((cantidad_09 - cantidad_08) / cantidad_08) * 100
                        if delta_pct > 0: color_delta, flecha = "#FF4B4B", "🔺"
                        elif delta_pct < 0: color_delta, flecha = "#09AB3B", "🔻"
                        else: color_delta, flecha = "#888888", "➖"
                        html_delta = f'<span style="color: {color_delta}; font-size: 24px; font-weight: bold;">{flecha} {abs(delta_pct):.1f}%</span> <span style="font-size: 20px;">vs Agosto</span>'

                    html_tarjeta = f"""
                    <div style="background-color: #ffffff; padding: 40px; border-radius: 16px; text-align: center; border: 1px solid #eee; border-top: 12px solid {colores_alerta[i]}; box-shadow: 4px 4px 15px rgba(0,0,0,0.05); height: 100%; animation: slideUpFade 0.6s ease-out {delay}s both;">
                        <p style="margin: 0; font-size: 40px; color: {colores_alerta[i]}; font-weight: 900;">#{i+1}</p>
                        <h3 style="margin: 10px 0 20px 0; color: #111; font-size: 32px; font-weight: 900;">{agente}</h3>
                        <div style="display: flex; justify-content: center; gap: 40px; margin-bottom: 20px;">
                            <div>
                                <p style="margin: 0; font-size: 20px; color: #666;">Septiembre</p>
                                <p style="margin: 0; font-size: 50px; font-weight: 900; color: #FF4B4B; animation: numberPop 0.5s ease-out {delay + 0.2}s both;">{cantidad_09}</p>
                            </div>
                            <div style="border-left: 2px solid #ddd; padding-left: 40px;">
                                <p style="margin: 0; font-size: 20px; color: #666;">Agosto</p>
                                <p style="margin: 0; font-size: 38px; font-weight: 600; color: #555; animation: numberPop 0.5s ease-out {delay + 0.3}s both;">{cantidad_08}</p>
                            </div>
                        </div>
                        <div style="background-color: #f8f9fa; border-radius: 8px; padding: 15px; border: 1px solid #eee;">
                            {html_delta}
                        </div>
                    </div>
                    """
                    with cols_agentes[i]:
                        st.markdown(html_tarjeta, unsafe_allow_html=True)
            else:
                st.warning(f"No se encontró la columna '{col_agente}' en ambos meses para comparar.")

        # ==========================================
        # DIAPOSITIVA 5: EVOLUCIÓN MENSUAL
        # ==========================================
        elif seleccion == "5. Evolución Mensual":
            mostrar_titulo_bloque("Evolución Mensual de Fallidas")
            
            datos_progreso = []
            porcentajes_fijos = {
                'ENERO': 14.69, 'FEBRERO': 15.45, 'MARZO': 15.10, 'ABRIL': 15.42,
                'MAYO': 17.75, 'JUNIO': 15.37, 'JULIO': 14.91, 'AGOSTO': 12.58, 'SEPTIEMBRE': 12.88,
            }

            for hoja in nombres_hojas:
                nombre_mes = mapa_meses.get(hoja, hoja) 
                if nombre_mes in porcentajes_fijos:
                    datos_progreso.append({"Mes": nombre_mes, "Mes_Original": nombre_mes, "Porcentaje": porcentajes_fijos[nombre_mes]})

            if datos_progreso:
                df_progreso = pd.DataFrame(datos_progreso)
                df_progreso['Mes_Original'] = pd.Categorical(df_progreso['Mes_Original'], categories=meses_ordenados, ordered=True)
                df_progreso = df_progreso.sort_values('Mes_Original')
                
                simbolos, colores_marcadores, colores_texto = ['circle'], ['#888888'], ['#555555']
                textos_html = [f"<span style='font-size:30px'><b>{df_progreso['Porcentaje'].iloc[0]:.2f}%</b></span><br><br><span style='font-size:22px'>Inicio</span><br>&nbsp;"]

                for i in range(1, len(df_progreso)):
                    pct_actual = df_progreso['Porcentaje'].iloc[i]
                    delta = pct_actual - df_progreso['Porcentaje'].iloc[i-1]
                    
                    if delta < 0:
                        simbolos.append('triangle-down'); colores_marcadores.append('#09AB3B'); colores_texto.append('#09AB3B')
                        textos_html.append(f"<span style='font-size:30px'><b>{pct_actual:.2f}%</b></span><br><br><span style='font-size:22px'>▼ {abs(delta):.2f}%</span><br>&nbsp;")
                    elif delta > 0:
                        simbolos.append('triangle-up'); colores_marcadores.append('#FF4B4B'); colores_texto.append('#FF4B4B')
                        textos_html.append(f"<span style='font-size:30px'><b>{pct_actual:.2f}%</b></span><br><br><span style='font-size:22px'>▲ {delta:.2f}%</span><br>&nbsp;")
                    else:
                        simbolos.append('circle'); colores_marcadores.append('#888888'); colores_texto.append('#555555')
                        textos_html.append(f"<span style='font-size:30px'><b>{pct_actual:.2f}%</b></span><br><br><span style='font-size:22px'>➖ 0.00%</span><br>&nbsp;")

                fig_trend = go.Figure()
                fig_trend.add_trace(go.Scatter(
                    x=df_progreso["Mes"], y=df_progreso["Porcentaje"], mode='lines',
                    line=dict(color='#dee2e6', width=6, shape='spline'), fill='tozeroy',
                    fillcolor='rgba(222, 226, 230, 0.2)', hoverinfo='skip', showlegend=False
                ))
                fig_trend.add_trace(go.Scatter(
                    x=df_progreso["Mes"], y=df_progreso["Porcentaje"], mode='markers+text',
                    text=textos_html, textposition='top center', textfont=dict(color=colores_texto), 
                    marker=dict(symbol=simbolos, size=35, color=colores_marcadores, line=dict(color='#FFFFFF', width=3)),
                    hovertemplate="<b>%{x}</b><br>Fallidas: %{y}%<extra></extra>", showlegend=False
                ))

                rango_max = df_progreso["Porcentaje"].max() + 14.0 
                fig_trend.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=750, 
                    xaxis=dict(showgrid=False, tickfont=dict(size=18, color='black', family="Inter"), title=""), 
                    yaxis=dict(showgrid=True, gridcolor='#f0f0f0', ticksuffix="%", tickfont=dict(size=18, color='#666'), title="<b>% de Fallidas</b>", title_font=dict(size=20), range=[0, rango_max]),
                    margin=dict(t=100, b=40, l=40, r=40),
                    hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
                )
                
                st.markdown('<div style="background-color: white; padding: 20px; border-radius: 16px; box-shadow: 0px 4px 15px rgba(0,0,0,0.05); border: 1px solid #eee; animation: slideUpFade 0.7s ease-out 0.2s both;">', unsafe_allow_html=True)
                st.plotly_chart(fig_trend, use_container_width=True, config={'displayModeBar': False})
                st.markdown('</div>', unsafe_allow_html=True)

        # ==========================================
        # DIAPOSITIVA 6: ÓRDENES CANCELADAS POR CANAL
        # ==========================================
        elif seleccion == "6. Canceladas por Canal":
            mostrar_titulo_bloque("Análisis de Órdenes Canceladas por Canal")
            
            if df_cancel is not None:
                total_canceladas = len(df_cancel)
                if col_canal in df_cancel.columns:
                    canales_counts = df_cancel[col_canal].value_counts()
                    columnas_por_fila = 5
                    cols_canales = st.columns(columnas_por_fila)
                    
                    with cols_canales[0]:
                        st.markdown(f"""
                        <div style="background-color: #ffffff; padding: 25px; border-radius: 12px; text-align: center; border: 1px solid #e0e0e0; border-top: 8px solid #343a40; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); height: 100%; margin-bottom: 20px; animation: slideUpFade 0.6s ease-out 0.1s both;">
                            <h3 style="margin: 0 0 15px 0; color: #555; font-size: 20px; font-weight: bold; text-transform: uppercase;">Total</h3>
                            <p style="margin: 0; font-size: 45px; color: #111; font-weight: 900; animation: numberPop 0.5s ease-out 0.25s both;">{total_canceladas}</p>
                            <p style="margin: 5px 0 0 0; font-size: 20px; color: #666; font-weight: bold;">12,37%</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    for i, (canal, cantidad) in enumerate(canales_counts.items()):
                        col_idx = (i + 1) % columnas_por_fila 
                        pct_canal = (cantidad / total_canceladas) * 100
                        delay = (i + 2) * 0.1
                        
                        html_canal = f"""
                        <div style="background-color: #ffffff; padding: 25px; border-radius: 12px; text-align: center; border: 1px solid #e0e0e0; border-top: 8px solid #FF4B4B; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); height: 100%; margin-bottom: 20px; animation: slideUpFade 0.6s ease-out {delay}s both;">
                            <h3 style="margin: 0 0 15px 0; color: #444; font-size: 20px; font-weight: bold;">{canal}</h3>
                            <p style="margin: 0; font-size: 45px; color: #FF4B4B; font-weight: 900; animation: numberPop 0.5s ease-out {delay + 0.15}s both;">{cantidad}</p>
                            <p style="margin: 5px 0 0 0; font-size: 16px; color: #888; font-weight: bold;">{pct_canal:.1f}%</p>
                        </div>
                        """
                        with cols_canales[col_idx]:
                            st.markdown(html_canal, unsafe_allow_html=True)
                else:
                    st.warning("No se encontró la columna 'CANAL' en la hoja 'Cancel'.")
            else:
                st.error("No se encontró la hoja 'Cancel' en el archivo Excel o no pudo ser leída correctamente.")

        # ==========================================
        # DIAPOSITIVA 7: MOTIVOS DE CANCELACIÓN (BARRAS VERTICALES)
        # ==========================================
        elif seleccion == "7. Motivos Cancelación":
            mostrar_titulo_bloque("Top Motivos de Cancelación")
            
            if df_cancel is not None:
                total_canceladas = len(df_cancel)
                if col_motivo in df_cancel.columns:
                    df_motivos_cancel = df_cancel[col_motivo].value_counts().reset_index()
                    df_motivos_cancel.columns = ['MOTIVO', 'CANTIDAD']
                    df_motivos_cancel = df_motivos_cancel.head(3).copy()
                    
                    df_motivos_cancel['PORCENTAJE'] = (df_motivos_cancel['CANTIDAD'] / total_canceladas) * 100
                    df_motivos_cancel['TEXTO_BARRA'] = df_motivos_cancel.apply(
                        lambda row: f"<span style='font-size:45px; font-weight:900; color:black;'>{row['PORCENTAJE']:.1f}%</span><br><br><span style='font-size:24px; color:black;'>({int(row['CANTIDAD'])})</span>", axis=1
                    )
                    colores_barras = ['#FF4B4B', '#fd7e14', '#ffc107']
                    
                    fig_motivos = go.Figure(go.Bar(
                        x=df_motivos_cancel['MOTIVO'],
                        y=df_motivos_cancel['PORCENTAJE'],
                        text=df_motivos_cancel['TEXTO_BARRA'],
                        textposition='outside',
                        textfont=dict(color='black'),
                        marker_color=colores_barras[:len(df_motivos_cancel)],
                        cliponaxis=False,
                        customdata=df_motivos_cancel['CANTIDAD'],
                        hovertemplate="<b>%{x}</b><br>Cantidad: %{customdata}<br>Representa el %{y:.1f}% del total general<extra></extra>"
                    ))
                    
                    max_y_cancel = df_motivos_cancel['PORCENTAJE'].max() if not df_motivos_cancel.empty else 100
                    fig_motivos.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=650,
                        xaxis=dict(showgrid=False, title="", tickfont=dict(size=18, color='black', family="Inter")),
                        yaxis=dict(showgrid=True, gridcolor='#f0f0f0', title="<b>% sobre el Total de Canceladas</b>", tickfont=dict(size=16, color='black'), range=[0, max_y_cancel * 1.5]), 
                        margin=dict(l=20, r=20, t=60, b=20),
                        hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
                    )
                    
                    st.markdown('<div style="background-color: white; padding: 20px; border-radius: 16px; box-shadow: 0px 4px 15px rgba(0,0,0,0.05); border: 1px solid #eee; animation: slideUpFade 0.7s ease-out 0.2s both;">', unsafe_allow_html=True)
                    st.plotly_chart(fig_motivos, use_container_width=True, config={'displayModeBar': False})
                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.warning("No se encontró la columna 'MOTIVO CANCELACION' en la hoja 'Cancel'.")
            else:
                st.error("No se encontró la hoja 'Cancel' en el archivo Excel o no pudo ser leída correctamente.")

        # ==========================================
        # DIAPOSITIVA 8: CLIENTE DESISTE DEL SERVICIO
        # ==========================================
        elif seleccion == "8. Desiste del Servicio":
            mostrar_titulo_bloque("Detalle: Cliente Desiste del Servicio")
            
            if df_cancel is not None:
                if col_motivo in df_cancel.columns and col_obs in df_cancel.columns:
                    df_desiste = df_cancel[df_cancel[col_motivo].astype(str).str.contains('CLIENTE DESISTE', na=False, case=False)].copy()
                    df_desiste = df_desiste.dropna(subset=[col_obs])
                    
                    basura = ['', 'NAN', 'NAT', 'NONE', 'NULL', '<NA>', 'N/A', 'NA', '-', '0', 'SIN REGISTRO WFM', 'SIN REGISTRO']
                    df_obs_validas = df_desiste[~df_desiste[col_obs].isin(basura)]
                    total_validas = len(df_obs_validas)
                    
                    if total_validas > 0:
                        st.markdown(f"""
                        <div style="background-color: #ffffff; padding: 20px; border-radius: 12px; border-left: 8px solid #FF4B4B; margin-bottom: 30px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); animation: slideUpFade 0.5s ease-out both;">
                        <p style="margin: 0; font-size: 18px; color: #555;">Total de observaciones válidas analizadas para "Cliente Desiste": <strong style="font-size: 24px; color: #111; animation: numberPop 0.5s ease-out 0.2s both; display: inline-block;">{total_validas}</strong></p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        df_obs_counts = df_obs_validas[col_obs].value_counts().reset_index()
                        df_obs_counts.columns = ['OBSERVACION', 'CANTIDAD']
                        
                        df_obs_counts['PORCENTAJE'] = (df_obs_counts['CANTIDAD'] / total_validas) * 100
                        df_obs_counts = df_obs_counts.sort_values('PORCENTAJE', ascending=True)
                        
                        df_obs_counts['TEXTO_BARRA'] = df_obs_counts.apply(
                            lambda row: f"<span style='font-size:28px; font-weight:900; color:black;'>{row['PORCENTAJE']:.1f}%</span>", axis=1
                        )
                        colores_barras = ['#FF4B4B'] * len(df_obs_counts)
                        
                        fig_obs = go.Figure(go.Bar(
                            x=df_obs_counts['PORCENTAJE'],
                            y=df_obs_counts['OBSERVACION'],
                            orientation='h',
                            text=df_obs_counts['TEXTO_BARRA'],
                            textposition='outside',
                            textfont=dict(color='black'),
                            marker_color=colores_barras,
                            cliponaxis=False,
                            hovertemplate="<b>%{y}</b><br>Representa el %{x:.1f}% de las observaciones<extra></extra>"
                        ))
                        
                        max_x_obs = df_obs_counts['PORCENTAJE'].max() if not df_obs_counts.empty else 100
                        altura_dinamica = max(400, len(df_obs_counts) * 70)
                        
                        fig_obs.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=altura_dinamica,
                            xaxis=dict(showgrid=False, showticklabels=False, title="", range=[0, max_x_obs * 1.25]),
                            yaxis=dict(showgrid=False, title="", tickfont=dict(size=14, color='black', family="Inter")), 
                            margin=dict(l=20, r=20, t=20, b=20),
                            hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
                        )
                        
                        st.markdown('<div style="background-color: white; padding: 20px; border-radius: 16px; box-shadow: 0px 4px 15px rgba(0,0,0,0.05); border: 1px solid #eee; animation: slideUpFade 0.7s ease-out 0.2s both;">', unsafe_allow_html=True)
                        st.plotly_chart(fig_obs, use_container_width=True, config={'displayModeBar': False})
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.info("No hay observaciones válidas para mostrar luego de excluir celdas vacías y 'SIN REGISTRO WFM'.")
                else:
                    st.warning("No se encontraron las columnas 'MOTIVO CANCELACION' o 'OBS' en la hoja 'Cancel'.")
            else:
                st.error("No se encontró la hoja 'Cancel' en el archivo Excel o no pudo ser leída correctamente.")

        # ==========================================
        # DIAPOSITIVA 9: CANAL AUTOGESTIÓN 
        # ==========================================
        elif seleccion == "9. Autogestión":
            mostrar_titulo_bloque("Detalle: Canal Autogestión (Duplicidad de Carga)")
            
            if df_cancel is not None:
                if col_canal in df_cancel.columns and col_obs in df_cancel.columns and col_motivo in df_cancel.columns:
                    condicion_canal = df_cancel[col_canal].astype(str).str.contains('AUTOGESTION', na=False, case=False)
                    condicion_motivo = df_cancel[col_motivo].astype(str).str.contains('DUPLICIDAD DE CARGA', na=False, case=False)
                    
                    df_auto = df_cancel[condicion_canal & condicion_motivo].copy()
                    df_auto = df_auto.dropna(subset=[col_obs])
                    
                    basura = ['', 'NAN', 'NAT', 'NONE', 'NULL', '<NA>', 'N/A', 'NA', '-', '0', 'SIN REGISTRO WFM', 'SIN REGISTRO']
                    df_obs_auto = df_auto[~df_auto[col_obs].isin(basura)]
                    total_auto = len(df_obs_auto)
                    
                    if total_auto > 0:
                        st.markdown(f"""
                        <div style="background-color: #ffffff; padding: 20px; border-radius: 12px; border-left: 8px solid #FF4B4B; margin-bottom: 30px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); animation: slideUpFade 0.5s ease-out both;">
                        <p style="margin: 0; font-size: 18px; color: #555;">Total de observaciones válidas analizadas en Autogestión: <strong style="font-size: 24px; color: #111; animation: numberPop 0.5s ease-out 0.2s both; display: inline-block;">{total_auto}</strong></p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        condicion_venta = df_obs_auto[col_obs].str.contains('VENTA PREVIA', na=False, case=False)
                        cant_venta_previa = condicion_venta.sum()
                        cant_otros = total_auto - cant_venta_previa
                        
                        df_pie_auto = pd.DataFrame({
                            'CATEGORIA': ['VENTA PREVIA', 'OTROS'],
                            'CANTIDAD': [cant_venta_previa, cant_otros]
                        })
                        colores_torta_auto = ['#FF4B4B', '#e0e0e0']
                        
                        fig_auto = go.Figure(data=[go.Pie(
                            labels=df_pie_auto['CATEGORIA'], 
                            values=df_pie_auto['CANTIDAD'],
                            textinfo='percent', 
                            textfont=dict(size=40, color='white', family="Inter"),
                            marker=dict(colors=colores_torta_auto, line=dict(color='#FFFFFF', width=4)),
                            hole=0.45, 
                            hovertemplate="<b>%{label}</b><br>Representa el %{percent} de Autogestión<extra></extra>"
                        )])
                        
                        fig_auto.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=600,
                            showlegend=True, 
                            legend=dict(font=dict(size=20, color='black'), orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5),
                            margin=dict(l=20, r=20, t=30, b=20),
                            hoverlabel=dict(bgcolor="#ffffff", font_size=15, font_family="Inter", font_color="#333333", bordercolor="#EAEAEA")
                        )
                        fig_auto.update_traces(insidetextfont=dict(color='white'), outsidetextfont=dict(color='black'))
                        
                        st.markdown('<div style="background-color: white; padding: 20px; border-radius: 16px; box-shadow: 0px 4px 15px rgba(0,0,0,0.05); border: 1px solid #eee; animation: slideUpFade 0.7s ease-out 0.2s both;">', unsafe_allow_html=True)
                        st.plotly_chart(fig_auto, use_container_width=True, config={'displayModeBar': False})
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.info("No hay observaciones válidas para mostrar en el canal Autogestión con motivo Duplicidad de Carga.")
                else:
                    st.warning("No se encontraron las columnas necesarias en la hoja 'Cancel'.")
    else:
        st.error("No se encontraron las hojas 'Base 08' o 'Base 09' en el Excel para realizar la comparativa.")

except FileNotFoundError:
    st.error(f"⚠️ No se encontró el archivo '{ruta_excel}'. Asegúrate de subirlo a tu repositorio de GitHub en la misma carpeta que este script.")
except Exception as e:
    st.error(f"Hubo un error al procesar el archivo: {e}")