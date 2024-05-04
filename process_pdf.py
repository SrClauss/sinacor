import pdfplumber
import multiprocessing
import os
import argparse



def credito_debito(valor, simbolo):
    if valor == 0.0:
        return 0.0

    if simbolo == "C":
        return valor

    if simbolo == "D":
        return float(valor.replace(",", "."))*-1.0

def get_page(pdf, i):
       

    page= pdf.pages[i]
    
    
    table_negociacao_croped = page.crop((20, 250, page.width, 450))

    negocios_realizados_datas = table_negociacao_croped.extract_table({
        "vertical_strategy": "explicit",
        "horizontal_strategy": "lines",
        "explicit_vertical_lines": [35, 44,93,111,173, 199, 293, 328, 380, 440, 543, 560],
        
    })
    
    
    table_resumo_financeiro = page.crop((506, 470, page.width, 650))
    #table_resumo_financeiro.to_image(resolution=300).draw_vlines([ 515,546, 559])
    #table_resumo_financeiro.to_image(resolution=300).draw_hlines([ 480, 489, 498, 516, 525, 534, 543,  570, 579, 588, 598, 607.5, 616.5, 625, 639, 648])
    

    resumo_financeiro = table_resumo_financeiro.extract_table({
        "vertical_strategy": "explicit",
        "horizontal_strategy": "explicit",
        "explicit_vertical_lines": [515, 546, 559],
        "explicit_horizontal_lines": [480, 489, 498, 516, 525, 534, 543,  570, 579, 588, 598, 607.5, 616.5, 625, 639, 648]
    })
    

    #page.crop((429, 60, 562, 70)).to_image(resolution=300)
    nota_folha_pregao = page.crop((429, 60, 562, 70)).extract_table()




    #page.crop((130,77, 250, 86)).to_image(resolution=300)
    corretora = page.crop((130,77, 250, 86)).extract_text()
    for papel in negocios_realizados_datas:
        valor =  {
            "nota": nota_folha_pregao[0][0],
            "folha": nota_folha_pregao[0][1],
            "date": nota_folha_pregao[0][2],
            "q": papel[0],
            "negociacao": papel[1],
            "op": papel[2],
            "tipo_mercado": papel[3],
            "prazo": papel[4],
            "papel": papel[5],
            "obs": papel[6],
            "quantidade": papel[7],
            "preço_ajuste": papel[8],
            "valor_operacao": papel[9],
            "d/c": papel[10],
            "valor_liquido_operacoes":credito_debito(resumo_financeiro[0][0], resumo_financeiro[0][1]),
            "taxa_liquidacao": credito_debito(resumo_financeiro[1][0], resumo_financeiro[1][1]),
            "taxa_registro": credito_debito(resumo_financeiro[2][0], resumo_financeiro[2][1]),
            "total_cblc": credito_debito(resumo_financeiro[3][0], resumo_financeiro[3][1]),
            "taxa_termo_opcoes": credito_debito(resumo_financeiro[4][0], resumo_financeiro[4][1]),
            "taxa_ana": credito_debito(resumo_financeiro[5][0], resumo_financeiro[5][1]),
            "emolumentos": credito_debito(resumo_financeiro[6][0], resumo_financeiro[6][1]),
            "taxa_operacional": credito_debito(resumo_financeiro[7][0], resumo_financeiro[7][1]),
            "execucao": credito_debito(resumo_financeiro[8][0], resumo_financeiro[8][1]),
            "taxa_custodia": credito_debito(resumo_financeiro[9][0], resumo_financeiro[9][1]),
            "impostos": credito_debito(resumo_financeiro[10][0], resumo_financeiro[10][1]),
            "irrf": credito_debito(resumo_financeiro[11][0], resumo_financeiro[11][1]),
            "outros": credito_debito(resumo_financeiro[12][0], resumo_financeiro[12][1]),
            "total_custos_despesas": credito_debito(resumo_financeiro[13][0], resumo_financeiro[13][1]),
            "liquido": credito_debito(resumo_financeiro[14][0], resumo_financeiro[14][1])}

       
        yield valor
            

def process_file(file):
    pdf = pdfplumber.open(file)
    results = []
    
    for i in range(len(pdf.pages)):
        for element in get_page(pdf, i):
            if element["papel"] != "":
                element["arquivo"] = pdf.path.name
                results.append(element)
    
    return results

def get_all_files(path):
    files = [file for file in os.listdir(path) if file.endswith(".pdf")]
    
    with multiprocessing.Pool(processes=multiprocessing.cpu_count()) as pool:
        results = pool.map(process_file, [path + "\\" + file for file in files])
    
    for result in results:
        for element in result:
            yield element




if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Processa arquivos PDF.')
    parser.add_argument('path', type=str, help='O caminho para os arquivos PDF.')
    args = parser.parse_args()

    for element in get_all_files(args.path):
        print(element)