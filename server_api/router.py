from flask import Flask, send_file, render_template, request, jsonify
import oracledb
import pandas as pd
from datetime import datetime, date
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from flask_cors import CORS
from waitress import serve
from dotenv import load_dotenv
import os
import logging
from dotenv import load_dotenv

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(message)s')


load_dotenv('.env')

app = Flask(__name__)  # Configura a pasta estática
CORS(app, resources={r"/reportHilti": {"origins": "*"}})

# @app.route('/')
# def index():
#     return render_template('index.html')  # Renderiza o index.html

@app.route("/reportHilti", methods=['GET'])
def reportHilti():
    ref_ext = request.args.get('ref_ext')
    tipo_nota = request.args.get('tipo_nota')
    data_de = request.args.get('data_de')
    data_ate = request.args.get('data_ate')
    
    #VALIDAÇÕES
    if tipo_nota is None:
        return jsonify({"error": "Parâmetro 'tipo_nota' é obrigatório"}), 400

    # Tratamento do tipo_nota
    tipo_nota_list = []
    try:
        for tn in tipo_nota.split(','):
            if tn.strip():
                tipo_nota_list.append(int(tn.strip()))
        if not tipo_nota_list:
            return jsonify({"error": "Pelo menos um tipo de nota válido deve ser fornecido"}), 400
        tipo_nota_str = ','.join(map(str, tipo_nota_list))
    except ValueError:
        return jsonify({"error": "Formato inválido para 'tipo_nota'. Use números inteiros separados por vírgula."}), 400

    print("Tipos de nota processados:", tipo_nota_list)
    print("String de tipos de nota para SQL:", tipo_nota_str)

    if not ref_ext or ref_ext.lower() == 'null':
        ref_ext_condition = "IS NOT NULL"
        ref_ext_param = {}
    else:
        ref_ext_list = ref_ext.split(',')
        ref_ext_condition = "IN ({})".format(','.join([':ref_ext{}'.format(i) for i in range(len(ref_ext_list))]))
        ref_ext_param = {f'ref_ext{i}': val.strip() for i, val in enumerate(ref_ext_list)}


    if data_de is None:
        hoje = date.today()
        data_de = date(hoje.year, hoje.month, 1).strftime('%d/%m/%Y')
    if data_ate is None:
        data_ate = datetime.now().strftime('%d/%m/%Y')

    print("Parâmetros recebidos:", ref_ext, tipo_nota, data_de, data_ate)

    if not all([tipo_nota]):
        return jsonify({"error": "Parâmetros incompletos"}), 400
    #VALIDAÇÕES 

    try:
        connection = oracledb.connect(
            user=os.getenv('USER'),
            password=os.getenv('PASSWORD'),
            host=os.getenv('HOST'),
            port=os.getenv('PORT'),
            service_name = os.getenv('SERVICE_NAME')
        )
        
        cursor = connection.cursor()

        # Preparar os tipos de nota
        tipo_nota_list = tipo_nota.split(',') if tipo_nota else ['0', '1', '3', '5', '9']
        tipo_nota_str = ','.join(tipo_nota_list)

        sql = """
        SELECT 
            PRC.PRI_COD AS REF_CONEXOS,
            PRC.PRI_ESP_REFCLIENTE AS REF_EXTERNA, 
            PRC.PRI_ESP_REFERENCIA AS REF_CLIENTE,
            INVITE.INI_ESP_NUMEROPO AS N_PO,
            TO_CHAR(SUBSTR(ITE.CDI_COD, 1, 2), 'FM00') || '/' || TO_CHAR(SUBSTR(ITE.CDI_COD, 3, 7), 'FM0000000') || '-' || TO_CHAR(SUBSTR(ITE.CDI_COD, 10, 1), 'FMX') AS DI,
            ADI.ADI_FLT_TAXA_DOLAR AS PTAX,
            DP.DPE_NOM_PESSOA AS CLIENTE_DO_PROCESO,
            COALESCE(P.PRD_ESP_COD_EXTERNO, '-') AS COD_EXTERNO,
            ITE.IDI_ESP_DESCR_COMPL AS DESCRICAO_PRODUTO,
            ADI.TEC_ESP_COD AS NCM,
            ITE.ADI_COD AS N_ADICAO,
            COALESCE(P.PRD_ESP_COD_EXPORT, '1051') AS COD_EXPORTADOR,
            (SELECT MAX(FAB.DPE_NOM_PESSOA) FROM CMN_DADOS_PESSOAS FAB WHERE FAB.PES_COD = ADI.ADI_COD_FABRICANTE) AS NOME_FABRICANTE,
            ITE.IDI_PRE_QUANTIDADE AS QUANTIDADE,
            CDP.DPR_FLT_PESO_LIQUIDO AS PESO_LIQUIDO,
            CDP.CFO_ESP_COD AS CFOP,
            CN.UND_DES_NOME AS UNIDADE,
            FD.DOC_ESP_NUMERO AS N_NOTA_FISCAL,
            ROUND(TRIB_FOB.DTR_MNY_VALORMN, 2) AS FOB,
            ROUND(TRIB_FRETE.DTR_MNY_VALORMN, 2) AS FRETE,
            ROUND(TRIB_SEGURO.DTR_MNY_VALORMN, 2) AS SEGURO,
            ROUND(TRIB_THC.DTR_MNY_VALORMN, 2) AS THC,
            ROUND((TRIB_FOB.DTR_MNY_VALORMN) + (TRIB_FRETE.DTR_MNY_VALORMN) + (TRIB_SEGURO.DTR_MNY_VALORMN) + (TRIB_THC.DTR_MNY_VALORMN), 2) AS CIF,
            ROUND(TRIB_II.DTR_MNY_VALORMN, 2) AS II,
            CASE
                WHEN NVL(TRIB_II.DTR_PCT_ALIQUOTA, 0) = 0 THEN ADI.ADI_PCT_ALIQ_II_XML
                ELSE TRIB_II.DTR_PCT_ALIQUOTA
            END AS PCT_II,
            NVL(TRIB_IPI.DTR_MNY_VALORMN, 0) AS IPI,
            NVL(TRIB_IPI.DTR_PCT_ALIQUOTA, 0) AS PCT_IPI,
            ROUND((TRIB_SISCOMEX.DTR_MNY_VALORMN + TRIB_AFRMM.DTR_MNY_VALORMN + TRIB_DESPESAS.DTR_MNY_VALORMN), 2) AS DESPESAS,
            NVL(ROUND(TRIB_PIS.DTR_MNY_VALORMN, 2), 0) AS PIS,
            NVL(
                CASE WHEN NVL(TRIB_PIS.DTR_PCT_ALIQUOTA, 0) = 0 THEN ADI.ADI_PCT_PIS
                ELSE TRIB_PIS.DTR_PCT_ALIQUOTA
            END, 0) AS PCT_PIS,
            NVL(ROUND(TRIB_COFINS.DTR_MNY_VALORMN, 2), 0) AS COFINS,
            NVL(
                CASE WHEN NVL(TRIB_COFINS.DTR_PCT_ALIQUOTA, 0) = 0 THEN ADI.ADI_PCT_COFINS
                ELSE TRIB_COFINS.DTR_PCT_ALIQUOTA
            END, 0) AS PCT_COFINS,
            NVL(ROUND(TRIB_ANTIDUMPING.DTR_MNY_VALORMN, 2), 0) AS ANTIDUMPING,
            NVL(ROUND(TRIB_BCICMS.DTR_MNY_VALORMN, 2), 0) AS BCICMS,
            NVL(ROUND(TRIB_ICMS.DTR_MNY_VALORMN, 2), 0) AS ICMS,
            NVL(TRIB_ICMS.DTR_PCT_ALIQUOTA, 0) AS PCT_ICMS,
            CDP.DPR_PRE_VALORUN AS VALOR_UNITARIO,
            NVL(ROUND((TRIB_IPI.DTR_MNY_VALORMN - TRIB_IPINAC.DTR_MNY_VALORMN), 2), 0) AS DIF_IPI,
            NVL(ROUND(TRIB_ICMSST.DTR_MNY_VALORMN, 2), 0) AS ICMS_ST,
            
            CASE 
                WHEN NVL(TRIB_ICMSST.DTR_MNY_VALORMN, 0) = 0 THEN TRIB_BCICMS.DTR_MNY_VALORMN
                ELSE TRIB_ICMSST.DTR_MNY_VALORMN + TRIB_BCICMS.DTR_MNY_VALORMN
            END AS TOTAL,
            CASE
                WHEN FD.DOC_VLD_TIPO = 0 THEN 'COMUM'
                WHEN FD.DOC_VLD_TIPO = 1 THEN 'RETORNO'
                WHEN FD.DOC_VLD_TIPO = 3 THEN 'COMPLEMENTAR'
                WHEN FD.DOC_VLD_TIPO = 5 THEN 'CORRETIVO'
                WHEN FD.DOC_VLD_TIPO = 9 THEN 'OUTROS DOCUMENTOS'
            END AS TIPO_NF
        FROM IMP_PROCESSO PRC
        JOIN IMP_DI_CABECALHO CAB ON PRC.FIL_COD = CAB.FIL_COD AND PRC.PRI_COD = CAB.PRI_COD 
        JOIN IMP_DI_ADICAO ADI ON CAB.FIL_COD = ADI.FIL_COD
            AND CAB.CDI_COD = ADI.CDI_COD
            AND CAB.CDI_COD_SEQ = ADI.CDI_COD_SEQ
        JOIN IMP_DI_ITEM ITE ON ADI.FIL_COD = ITE.FIL_COD AND ADI.CDI_COD = ITE.CDI_COD AND ADI.CDI_COD_SEQ = ITE.CDI_COD_SEQ AND ADI.ADI_COD = ITE.ADI_COD
        JOIN COM_PRODUTOS P ON ITE.PRD_COD = P.PRD_COD
        JOIN IMP_DI_ITEM_INV ITEINV ON ITE.FIL_COD = ITEINV.FIL_COD AND ITE.CDI_COD = ITEINV.CDI_COD AND ITE.CDI_COD_SEQ = ITEINV.CDI_COD_SEQ AND ITE.ADI_COD = ITEINV.ADI_COD AND ITE.IDI_COD = ITEINV.IDI_COD AND ITE.PRD_COD = ITEINV.PRD_COD
        JOIN PRC_INVOICE_ITENS INVITE ON ITEINV.INV_COD = INVITE.INV_COD AND ITEINV.INI_ITEM = INVITE.INI_ITEM
        JOIN PRC_INVOICE INV ON INVITE.INV_COD = INV.INV_COD
        JOIN COM_DOC_PRODUTOS CDP ON ITE.FIL_COD = CDP.FIL_COD AND ITE.CDI_COD = CDP.CDI_COD AND ITE.CDI_COD_SEQ = CDP.CDI_COD_SEQ AND ITE.ADI_COD = CDP.ADI_COD AND ITE.IDI_COD = CDP.IDI_COD
        JOIN FIN_DOC FD ON CDP.FIL_COD = FD.FIL_COD AND CDP.DOC_TIP = FD.DOC_TIP AND CDP.DOC_COD = FD.DOC_COD
        JOIN CMN_DADOS_PESSOAS DP ON FD.PES_COD = DP.PES_COD AND FD.DPE_COD_SEQ = DP.DPE_COD_SEQ
        JOIN CMN_PESSOAS CP ON DP.PES_COD = CP.PES_COD
        JOIN COM_PROD_UNIDADES PU ON P.PRD_COD = PU.PRD_COD AND P.UND_COD = PU.UND_COD
        JOIN COM_UNIDADES CN ON P.UND_COD = CN.UND_COD
        LEFT JOIN COM_DOC_PRODTRIB TRIB_II ON CDP.FIL_COD = TRIB_II.FIL_COD AND CDP.DOC_TIP = TRIB_II.DOC_TIP AND CDP.DOC_COD = TRIB_II.DOC_COD AND CDP.FIS_COD = TRIB_II.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_II.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_II.PRD_COD AND TRIB_II.IMP_COD = 630
        LEFT JOIN COM_DOC_PRODTRIB TRIB_IPI ON CDP.FIL_COD = TRIB_IPI.FIL_COD AND CDP.DOC_TIP = TRIB_IPI.DOC_TIP AND CDP.DOC_COD = TRIB_IPI.DOC_COD AND CDP.FIS_COD = TRIB_IPI.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_IPI.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_IPI.PRD_COD AND TRIB_IPI.IMP_COD = 8
        LEFT JOIN COM_DOC_PRODTRIB TRIB_PIS ON CDP.FIL_COD = TRIB_PIS.FIL_COD AND CDP.DOC_TIP = TRIB_PIS.DOC_TIP AND CDP.DOC_COD = TRIB_PIS.DOC_COD AND CDP.FIS_COD = TRIB_PIS.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_PIS.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_PIS.PRD_COD AND TRIB_PIS.IMP_COD = 939
        LEFT JOIN COM_DOC_PRODTRIB TRIB_COFINS ON CDP.FIL_COD = TRIB_COFINS.FIL_COD AND CDP.DOC_TIP = TRIB_COFINS.DOC_TIP AND CDP.DOC_COD = TRIB_COFINS.DOC_COD AND CDP.FIS_COD = TRIB_COFINS.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_COFINS.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_COFINS.PRD_COD AND TRIB_COFINS.IMP_COD = 940
        LEFT JOIN COM_DOC_PRODTRIB TRIB_ICMS ON CDP.FIL_COD = TRIB_ICMS.FIL_COD AND CDP.DOC_TIP = TRIB_ICMS.DOC_TIP AND CDP.DOC_COD = TRIB_ICMS.DOC_COD AND CDP.FIS_COD = TRIB_ICMS.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_ICMS.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_ICMS.PRD_COD AND TRIB_ICMS.IMP_COD = 7
        LEFT JOIN COM_DOC_PRODTRIB TRIB_FOB ON CDP.FIL_COD = TRIB_FOB.FIL_COD AND CDP.DOC_TIP = TRIB_FOB.DOC_TIP AND CDP.DOC_COD = TRIB_FOB.DOC_COD AND CDP.FIS_COD = TRIB_FOB.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_FOB.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_FOB.PRD_COD AND TRIB_FOB.IMP_COD = 600
        LEFT JOIN COM_DOC_PRODTRIB TRIB_FRETE ON CDP.FIL_COD = TRIB_FRETE.FIL_COD AND CDP.DOC_TIP = TRIB_FRETE.DOC_TIP AND CDP.DOC_COD = TRIB_FRETE.DOC_COD AND CDP.FIS_COD = TRIB_FRETE.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_FRETE.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_FRETE.PRD_COD AND TRIB_FRETE.IMP_COD = 610
        LEFT JOIN COM_DOC_PRODTRIB TRIB_SEGURO ON CDP.FIL_COD = TRIB_SEGURO.FIL_COD AND CDP.DOC_TIP = TRIB_SEGURO.DOC_TIP AND CDP.DOC_COD = TRIB_SEGURO.DOC_COD AND CDP.FIS_COD = TRIB_SEGURO.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_SEGURO.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_SEGURO.PRD_COD AND TRIB_SEGURO.IMP_COD = 620
        LEFT JOIN COM_DOC_PRODTRIB TRIB_THC ON CDP.FIL_COD = TRIB_THC.FIL_COD AND CDP.DOC_TIP = TRIB_THC.DOC_TIP AND CDP.DOC_COD = TRIB_THC.DOC_COD AND CDP.FIS_COD = TRIB_THC.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_THC.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_THC.PRD_COD AND TRIB_THC.IMP_COD = 1000
        LEFT JOIN COM_DOC_PRODTRIB TRIB_SISCOMEX ON CDP.FIL_COD = TRIB_SISCOMEX.FIL_COD AND CDP.DOC_TIP = TRIB_SISCOMEX.DOC_TIP AND CDP.DOC_COD = TRIB_SISCOMEX.DOC_COD AND CDP.FIS_COD = TRIB_SISCOMEX.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_SISCOMEX.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_SISCOMEX.PRD_COD AND TRIB_SISCOMEX.IMP_COD = 983
        LEFT JOIN COM_DOC_PRODTRIB TRIB_AFRMM ON CDP.FIL_COD = TRIB_AFRMM.FIL_COD AND CDP.DOC_TIP = TRIB_AFRMM.DOC_TIP AND CDP.DOC_COD = TRIB_AFRMM.DOC_COD AND CDP.FIS_COD = TRIB_AFRMM.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_AFRMM.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_AFRMM.PRD_COD AND TRIB_AFRMM.IMP_COD = 1017
        LEFT JOIN COM_DOC_PRODTRIB TRIB_DESPESAS ON CDP.FIL_COD = TRIB_DESPESAS.FIL_COD AND CDP.DOC_TIP = TRIB_DESPESAS.DOC_TIP AND CDP.DOC_COD = TRIB_DESPESAS.DOC_COD AND CDP.FIS_COD = TRIB_DESPESAS.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_DESPESAS.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_DESPESAS.PRD_COD AND TRIB_DESPESAS.IMP_COD = 650
        LEFT JOIN COM_DOC_PRODTRIB TRIB_ANTIDUMPING ON CDP.FIL_COD = TRIB_ANTIDUMPING.FIL_COD AND CDP.DOC_TIP = TRIB_ANTIDUMPING.DOC_TIP AND CDP.DOC_COD = TRIB_ANTIDUMPING.DOC_COD AND CDP.FIS_COD = TRIB_ANTIDUMPING.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_ANTIDUMPING.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_ANTIDUMPING.PRD_COD AND TRIB_ANTIDUMPING.IMP_COD = 986
        LEFT JOIN COM_DOC_PRODTRIB TRIB_BCICMS ON CDP.FIL_COD = TRIB_BCICMS.FIL_COD AND CDP.DOC_TIP = TRIB_BCICMS.DOC_TIP AND CDP.DOC_COD = TRIB_BCICMS.DOC_COD AND CDP.FIS_COD = TRIB_BCICMS.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_BCICMS.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_BCICMS.PRD_COD AND TRIB_BCICMS.IMP_COD = 4
        LEFT JOIN COM_DOC_PRODTRIB TRIB_ICMSST ON CDP.FIL_COD = TRIB_ICMSST.FIL_COD AND CDP.DOC_TIP = TRIB_ICMSST.DOC_TIP AND CDP.DOC_COD = TRIB_ICMSST.DOC_COD AND CDP.FIS_COD = TRIB_ICMSST.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_ICMSST.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_ICMSST.PRD_COD AND TRIB_ICMSST.IMP_COD = 9
        LEFT JOIN COM_DOC_PRODTRIB TRIB_IPINAC ON CDP.FIL_COD = TRIB_IPINAC.FIL_COD AND CDP.DOC_TIP = TRIB_IPINAC.DOC_TIP AND CDP.DOC_COD = TRIB_IPINAC.DOC_COD AND CDP.FIS_COD = TRIB_IPINAC.FIS_COD AND CDP.DPR_COD_SEQ = TRIB_IPINAC.DPR_COD_SEQ AND CDP.PRD_COD = TRIB_IPINAC.PRD_COD AND TRIB_IPINAC.IMP_COD = 640
        WHERE
            CAB.CDI_DTA_REGISTRO BETWEEN TO_DATE(:data_de, 'DD/MM/YYYY') AND TO_DATE(:data_ate, 'DD/MM/YYYY')
            AND PRC.PRI_ESP_REFCLIENTE {}
            AND FD.DOC_VLD_TIPO IN ({})
            AND FD.FIL_COD IN (1, 5, 12)
            AND FD.PES_COD = '22178'
            AND FD.DOC_VLD_FINALIZADO = 1
            AND FD.DOC_VLD_CANCELAMENTO = 0
        """.format(ref_ext_condition, tipo_nota_str)

        params = {
            'data_de': data_de,
            'data_ate': data_ate,
            **ref_ext_param
        }

        print("Parâmetros para SQL:", params)

        cursor.execute(sql, params)
        results = cursor.fetchall()

        print(f"Número de resultados: {len(results)}")

        if not results:
            print("Nenhum resultado encontrado")
            return jsonify({"message": "Nenhum resultado encontrado"}), 404

        df = pd.DataFrame(results, columns=[desc[0] for desc in cursor.description])
        
        output = io.BytesIO()
        wb = formatacao(df)
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='relatorio_hilti.xlsx'
        )

    except oracledb.DatabaseError as e:
        print(f"Erro de banco de dados: {e}")
        return jsonify({"error": f"Erro ao acessar o banco de dados: {str(e)}"}), 500
    except Exception as e:
        print(f"Erro inesperado: {str(e)}")
        return jsonify({"error": f"Ocorreu um erro inesperado: {str(e)}"}), 500
    finally:
        if 'connection' in locals() and connection is not None:
            connection.close()

def formatacao(df):
    wb = Workbook()
    ws_sql = wb.active
    ws_sql.title = "Rateio de produtos"

    # Função para escrever DataFrame em um sheet
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_sql.append(r)

    def format_sheet(ws):
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        blue_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        for idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
            if idx % 2 == 0:
                for cell in row:
                    cell.fill = blue_fill
            else:
                for cell in row:
                    cell.fill = white_fill

    # Aplica a formatação
    format_sheet(ws_sql)
    return wb

if __name__ == '__main__':
    # Configurações para produção
    host = '0.0.0.0'
    port = 5000
    threads = 2

    
    serve(app, host=host, port=port, threads=threads)

