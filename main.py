import pandas as pd
import fdb

excel_path1 = r'C:/Users/Gabriel/Desktop/Fluxo.xlsx'
dst_path = r'MTK:C:/Microsys/MsysIndustrial/Dados/MSYSDADOS.FDB'
TABLE_NAME1 = 'PEDIDOS_VENDAS_PARCELAS'
TABLE_NAME2 = 'PEDIDOS_VENDAS'
TABLE_NAME3 = 'PEDIDOS_VENDAS_ITENS'

########################################################################
SELECT1 = 'select PVP_PDV_NUMERO, PVP_VENCIMENTO, PVP_VALOR from %s ' \
            % (TABLE_NAME1)
SELECT2 = 'select PDV_NUMERO, PDV_DATA, PDV_CLI_CODIGO, PDV_PSI_CODIGO, PDV_CON_CODIGO, PDV_VALORPRODUTOS from %s ' \
            % (TABLE_NAME2)
SELECT3 = 'select PVI_NUMERO, PVI_QUANTIDADE, PVI_VL_CUSTO_ITEM from %s ' \
            % (TABLE_NAME3)
########################################################################
con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')
cur = con.cursor()
####################################################################################
cur.execute(SELECT1)
table_rows1 = cur.fetchall()
df1 = pd.DataFrame(table_rows1)
####################################################################################
cur.execute(SELECT2)
table_rows2 = cur.fetchall()
df2 = pd.DataFrame(table_rows2)
####################################################################################
cur.execute(SELECT3)
table_rows3 = cur.fetchall()
df3 = pd.DataFrame(table_rows3)

#PEDIDOS FILTRADOS POR DATA
df1[1] = pd.to_datetime(df1[1], format='%Y-%m-%d')
df1 = df1.groupby(0).apply(lambda x: x)
########################################################################
df2[1] = pd.to_datetime(df2[1], format='%Y-%m-%d')
df2 = df2[df2[4] != 36]
df2 = df2.drop(columns=4)
df2 = df2[df2[3] != "CC"]
df2 = df2[df2[3] != "AA"]
########################################################################
df3[3] = df3[1]*df3[2]
df3 = df3.drop(columns=1)
df3 = df3.drop(columns=2)


with pd.ExcelWriter(excel_path1) as writer:
    df1.to_excel(writer, index=False, sheet_name='PARCELAS')
    df2.to_excel(writer, index=False, sheet_name='PEDIDOS')
    df3.to_excel(writer, index=False, sheet_name='VALOR_ITENS')
