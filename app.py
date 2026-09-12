import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conciliador Bancario", layout="wide", page_icon="🏦")

st.title("🏦 Conciliador Bancario Automatizado")
st.markdown("Carga el **Mayor Contable**, el **Extracto Bancario** y (opcionalmente) las **Liquidaciones de Tarjetas**.")

# Sección de carga de archivos principales
st.subheader("1. Archivos Principales")
col1, col2 = st.columns(2)
with col1:
    file_mayor = st.file_uploader("Mayor Contable (.xlsx)", type=["xlsx"], key="mayor")
with col2:
    file_banco = st.file_uploader("Extracto Bancario (.xlsx)", type=["xlsx"], key="banco")

# Sección opcional de tarjetas
st.subheader("2. Liquidaciones de Tarjetas (Opcional)")
st.info("Si adjuntas las liquidaciones, la app desglosará automáticamente los cobros con tarjeta sin que tengas que agregarlos a mano al Mayor.")

col_t1, col_t2, col_t3, col_t4 = st.columns(4)
with col_t1:
    file_fiserv = st.file_uploader("Liquidación Fiserv (.xlsx)", type=["xlsx"], key="fiserv")
with col_t2:
    file_cabal = st.file_uploader("Liquidación Cabal (.xlsx)", type=["xlsx"], key="cabal")
with col_t3:
    file_confiable = st.file_uploader("Liquidación Confiable (.xlsx)", type=["xlsx"], key="confiable")
with col_t4:
    file_qr = st.file_uploader("Liquidación QR (.xlsx)", type=["xlsx"], key="qr")

# Parámetros de tolerancia
st.sidebar.header("⚙️ Parámetros de Conciliación")
tol_dias = st.sidebar.number_input("Tolerancia Días (±)", min_value=0, max_value=60, value=15)
tol_monto = st.sidebar.number_input("Tolerancia Monto (±$)", min_value=0.00, max_value=10.00, value=0.05, step=0.01)

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

def parse_card_files(f_fiserv, f_cabal, f_confiable, f_qr):
    cards_dfs = []
    
    if f_fiserv:
        df = pd.read_excel(f_fiserv).dropna(subset=['FECHA DE PAGO']).copy()
        df['Fecha_dt'] = pd.to_datetime(df['FECHA DE PAGO'], format='%d/%m/%Y', errors='coerce')
        df['Neto'] = pd.to_numeric(df['IMPORTE NETO'], errors='coerce')
        df['Descripción'] = "Fiserv - " + df['TARJETA'].astype(str)
        df['Origen'] = 'Fiserv'
        cards_dfs.append(df[['Fecha_dt', 'Neto', 'Descripción', 'Origen']])

    if f_cabal:
        df = pd.read_excel(f_cabal).dropna(subset=['Fecha de pago']).copy()
        df['Fecha_dt'] = pd.to_datetime(df['Fecha de pago'], format='%d/%m/%Y', errors='coerce')
        df['Neto'] = pd.to_numeric(df['Importe neto final a liquidar'], errors='coerce')
        df['Descripción'] = "Liquidación Cabal"
        df['Origen'] = 'Cabal'
        cards_dfs.append(df[['Fecha_dt', 'Neto', 'Descripción', 'Origen']])

    if f_confiable:
        df = pd.read_excel(f_confiable).dropna(subset=['Fecha de pago']).copy()
        df['Fecha_dt'] = pd.to_datetime(df['Fecha de pago'], format='%d/%m/%Y', errors='coerce')
        df['Neto'] = pd.to_numeric(df['Importe neto final a liquidar'], errors='coerce')
        df['Descripción'] = "Liquidación Confiable"
        df['Origen'] = 'Confiable'
        cards_dfs.append(df[['Fecha_dt', 'Neto', 'Descripción', 'Origen']])

    if f_qr:
        df = pd.read_excel(f_qr).dropna(subset=['Fecha Operación']).copy()
        df['Fecha_dt'] = pd.to_datetime(df['Fecha Operación'], format='%d/%m/%Y', errors='coerce')
        df['Neto'] = pd.to_numeric(df['Monto acreditado'], errors='coerce')
        df['Descripción'] = "Liquidación QR - " + df['Transacción'].astype(str)
        df['Origen'] = 'QR'
        cards_dfs.append(df[['Fecha_dt', 'Neto', 'Descripción', 'Origen']])

    if cards_dfs:
        return pd.concat(cards_dfs, ignore_index=True)
    return pd.DataFrame()

if file_mayor and file_banco:
    # 1. Parsear Banco
    df_b = parse_banco(file_banco)
    
    # 2. Parsear Mayor
    df_m = pd.read_excel(file_mayor, sheet_name=0)
    if 'Documento' in df_m.columns:
        df_m = df_m[df_m['Documento'] != 'Saldo Inicial'].copy()
    
    has_card_files = any([file_fiserv, file_cabal, file_confiable, file_qr])
    
    if has_card_files:
        # Filtrar asientos sumarios globales de tarjetas si se adjuntan las liquidaciones
        df_m = df_m[~df_m['Descripción'].astype(str).str.contains('Liquidación|LIQTAR', case=False, na=False)].copy()
        df_m = df_m[~df_m['Documento'].astype(str).str.contains('SR-LIQTAR', case=False, na=False)].copy()

    df_m['Neto'] = df_m['Debe'].fillna(0) - df_m['Haber'].fillna(0)
    df_m['Fecha_dt'] = pd.to_datetime(df_m['Fecha'])
    df_m['Origen'] = 'Mayor'
    
    # 3. Parsear Tarjetas
    df_cards = parse_card_files(file_fiserv, file_cabal, file_confiable, file_qr)
    
    # Consolidar Registros Internos
    cols = ['Fecha_dt', 'Neto', 'Descripción', 'Origen']
    df_internal = pd.concat([df_m[cols], df_cards[cols]], ignore_index=True) if not df_cards.empty else df_m[cols]
    
    # Matching
    matched_banco = set()
    matched_internal = set()
    results = []

    for idx_i, row_i in df_internal.iterrows():
        i_amount = row_i['Neto']
        i_date = row_i['Fecha_dt']
        
        candidates = df_b[~df_b.index.isin(matched_banco)].copy()
        
        amt_diff = (candidates['Importe_num'] - i_amount).abs()
        date_diff = (candidates['Fecha_dt'] - i_date).abs().dt.days
        
        valid = candidates[(amt_diff <= tol_monto) & (date_diff <= tol_dias)]
        
        if len(valid) > 0:
            valid = valid.assign(
                diff_dias=(valid['Fecha_dt'] - i_date).abs().dt.days,
                diff_monto=(valid['Importe_num'] - i_amount).abs()
            ).sort_values(by=['diff_dias', 'diff_monto'])
            
            best_idx = valid.index[0]
            matched_internal.add(idx_i)
            matched_banco.add(best_idx)
            
            b_row = valid.loc[best_idx]
            results.append({
                'Estado': 'CONCILIADO',
                'Origen': row_i['Origen'],
                'Fecha_Interna': row_i['Fecha_dt'].strftime('%Y-%m-%d'),
                'Descripción_Interna': row_i['Descripción'],
                'Monto_Interno': row_i['Neto'],
                'Banco_Fecha': b_row['Fecha_dt'].strftime('%Y-%m-%d'),
                'Banco_Descripción': b_row.get('Descripción', ''),
                'Banco_Monto': b_row['Importe_num'],
                'Diferencia_Días': valid.loc[best_idx, 'diff_dias'],
                'Diferencia_Monto': valid.loc[best_idx, 'diff_monto']
            })

    df_conciliados = pd.DataFrame(results)
    df_unmatched_internal = df_internal[~df_internal.index.isin(matched_internal)].copy()
    df_unmatched_banco = df_b[~df_b.index.isin(matched_banco)].copy()
    
    st.success("✅ ¡Conciliación procesada con éxito!")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Movimientos Conciliados", len(df_conciliados))
    c2.metric("Pendientes Internos", len(df_unmatched_internal))
    c3.metric("Pendientes en Banco", len(df_unmatched_banco))
    
    tab1, tab2, tab3 = st.tabs(["🟢 Conciliados", "🔴 Pendientes Internos (Mayor/Tarjetas)", "🟡 Pendientes en Banco"])
    
    with tab1:
        st.dataframe(df_conciliados, use_container_width=True)
    with tab2:
        st.dataframe(df_unmatched_internal, use_container_width=True)
    with tab3:
        st.dataframe(df_unmatched_banco[['Fecha', 'Concepto/Cod.Op.', 'Descripción', 'Importe_num']], use_container_width=True)
        
    output_filename = "Conciliacion_Consolidada_Resultado.xlsx"
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
        df_unmatched_internal.to_excel(writer, sheet_name='Pendientes_Internos', index=False)
        df_unmatched_banco.to_excel(writer, sheet_name='Pendientes_Banco', index=False)
        
    with open(output_filename, "rb") as f:
        st.download_button(
            label="📥 Descargar Reporte Completo en Excel",
            data=f,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
