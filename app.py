import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conciliador Bancario", layout="wide", page_icon="🏦")

st.title("🏦 Conciliación Bancaria Automatizada")
st.markdown("Arrastra el **Mayor Contable** y el **Resumen Bancario** para realizar el cruzamiento.")

col1, col2 = st.columns(2)
with col1:
    file_mayor = st.file_uploader("1. Subir Mayor Contable (.xlsx)", type=["xlsx"])
with col2:
    file_banco = st.file_uploader("2. Subir Resumen Bancario (.xlsx)", type=["xlsx"])

# Configuración en la barra lateral
st.sidebar.header("⚙️ Parámetros de Conciliación")
tol_dias = st.sidebar.number_input("Tolerancia Días (±)", min_value=0, max_value=60, value=15)
tol_monto = st.sidebar.number_input("Tolerancia Monto (±$)", min_value=0.00, max_value=10.00, value=0.05, step=0.01)

def parse_mayor(file):
    df = pd.read_excel(file, sheet_name=0)
    if 'Documento' in df.columns:
        df = df[df['Documento'] != 'Saldo Inicial'].copy()
    df['Neto'] = df['Debe'].fillna(0) - df['Haber'].fillna(0)
    df['Fecha_dt'] = pd.to_datetime(df['Fecha'])
    return df

def parse_banco(file):
    xls = pd.ExcelFile(file)
    sheet = 'principal' if 'principal' in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(file, sheet_name=sheet)
    
    if 'Fecha' not in df_raw.columns:
        header_idx = df_raw[df_raw.iloc[:, 0] == 'Concepto/Cod.Op.'].index
        if len(header_idx) > 0:
            idx = header_idx[0]
            df_raw.columns = df_raw.iloc[idx].values
            df_raw = df_raw.iloc[idx+1:].copy()
            
    df = df_raw.dropna(subset=['Fecha', 'Importe']).copy()
    df['Importe_num'] = pd.to_numeric(df['Importe'], errors='coerce')
    df['Fecha_dt'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y', errors='coerce')
    df = df.dropna(subset=['Importe_num', 'Fecha_dt']).copy()
    return df

if file_mayor and file_banco:
    df_m = parse_mayor(file_mayor)
    df_b = parse_banco(file_banco)
    
    matched_mayor = set()
    matched_banco = set()
    results = []

    for idx_m, row_m in df_m.iterrows():
        m_amount = row_m['Neto']
        m_date = row_m['Fecha_dt']
        
        candidates = df_b[~df_b.index.isin(matched_banco)].copy()
        
        amt_diff = (candidates['Importe_num'] - m_amount).abs()
        date_diff = (candidates['Fecha_dt'] - m_date).abs().dt.days
        
        valid = candidates[(amt_diff <= tol_monto) & (date_diff <= tol_dias)]
        
        if len(valid) > 0:
            valid = valid.assign(
                diff_dias=(valid['Fecha_dt'] - m_date).abs().dt.days,
                diff_monto=(valid['Importe_num'] - m_amount).abs()
            ).sort_values(by=['diff_dias', 'diff_monto'])
            
            best_idx = valid.index[0]
            matched_mayor.add(idx_m)
            matched_banco.add(best_idx)
            
            b_row = valid.loc[best_idx]
            results.append({
                'Estado': 'CONCILIADO',
                'Mayor_Fecha': row_m['Fecha_dt'].strftime('%Y-%m-%d'),
                'Mayor_Documento': row_m.get('Documento', ''),
                'Mayor_Descripción': row_m.get('Descripción', ''),
                'Mayor_Monto': row_m['Neto'],
                'Banco_Fecha': b_row['Fecha_dt'].strftime('%Y-%m-%d'),
                'Banco_Descripción': b_row.get('Descripción', ''),
                'Banco_Monto': b_row['Importe_num'],
                'Diferencia_Días': valid.loc[best_idx, 'diff_dias'],
                'Diferencia_Monto': valid.loc[best_idx, 'diff_monto']
            })

    df_conciliados = pd.DataFrame(results)
    df_unmatched_mayor = df_m[~df_m.index.isin(matched_mayor)].copy()
    df_unmatched_banco = df_b[~df_b.index.isin(matched_banco)].copy()
    
    st.success("✅ ¡Conciliación procesada con éxito!")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Movimientos Conciliados", len(df_conciliados))
    c2.metric("Pendientes en Mayor", len(df_unmatched_mayor))
    c3.metric("Pendientes en Banco", len(df_unmatched_banco))
    
    tab1, tab2, tab3 = st.tabs(["🟢 Conciliados", "🔴 Pendientes en Mayor", "🟡 Pendientes en Banco"])
    
    with tab1:
        st.dataframe(df_conciliados, use_container_width=True)
    with tab2:
        st.dataframe(df_unmatched_mayor[['Fecha', 'Documento', 'Descripción', 'Neto']], use_container_width=True)
    with tab3:
        st.dataframe(df_unmatched_banco[['Fecha', 'Concepto/Cod.Op.', 'Descripción', 'Importe_num']], use_container_width=True)
        
    output_filename = "Conciliacion_Resultado.xlsx"
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
        df_unmatched_mayor.to_excel(writer, sheet_name='Pendientes_Mayor', index=False)
        df_unmatched_banco.to_excel(writer, sheet_name='Pendientes_Banco', index=False)
        
    with open(output_filename, "rb") as f:
        st.download_button(
            label="📥 Descargar Reporte en Excel",
            data=f,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
