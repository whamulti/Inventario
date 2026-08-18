import os
import pdfplumber
import pandas as pd
import re
from datetime import datetime

# Cabeçalho de produto: "<código> - <nome> ... U ... * <UN|MT|RL|...>"
# O código pode ter poucos dígitos, letras (ex: "105C") ou pontos (ex: "50.0261"),
# então a detecção não pode depender da quantidade de dígitos.
PADRAO_CODIGO_NOME = re.compile(r'^(\S+?)\s*-\s*(.+)$')
# Terminador normal: "* UN", "* MT" etc.
PADRAO_FIM_UNIDADE = re.compile(r'\*\s*[A-Za-z]{2,4}\s*$')
# Quando a extração do PDF corta a linha em "... U *" (unidade perdida na quebra)
PADRAO_U_ISOLADO = re.compile(r'\bU\b')
# Quando a linha termina em "... 42 UN" (grade de tamanhos, sem asterisco)
PADRAO_FIM_UNIDADE_SEM_ASTERISCO = re.compile(r'\d\s+[A-Za-z]{2,4}\s*$')


def extrair_cabecalho_produto(linha_limpa):
    """Retorna (código, nome) se a linha for um cabeçalho de produto, senão None."""
    m = PADRAO_CODIGO_NOME.match(linha_limpa)
    if not m:
        return None
    codigo, resto = m.group(1), m.group(2).strip()
    if (PADRAO_FIM_UNIDADE.search(resto)
            or PADRAO_U_ISOLADO.search(resto)
            or PADRAO_FIM_UNIDADE_SEM_ASTERISCO.search(resto)):
        return codigo, resto
    return None

def extrair_produtos_inventario(caminho_pdf, arquivo_log):
    produtos = []
    paginas_com_reserva = []

    with open(arquivo_log, 'w', encoding='utf-8') as log:
        log.write("="*80 + "\n")
        log.write(f"LOG DE EXTRAÇÃO DE INVENTÁRIO\n")
        log.write(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        log.write(f"Arquivo: {caminho_pdf}\n")
        log.write("="*80 + "\n\n")

        with pdfplumber.open(caminho_pdf) as pdf:
            total_paginas = len(pdf.pages)
            log.write(f"Total de páginas no PDF: {total_paginas}\n\n")

            cod_produto = None
            nome_produto = None
            descricao_cor = None
            qtde_estoque = None
            qtde_reservada = None
            pagina_num = 0
            total_geral_estoque = None
            total_geral_reservada = None

            def flush_produto():
                nonlocal cod_produto, nome_produto, descricao_cor, qtde_estoque, qtde_reservada
                if cod_produto and qtde_estoque is not None:
                    reservada = qtde_reservada if qtde_reservada is not None else 0
                    produtos.append({
                        'Código': cod_produto,
                        'Nome do Produto': nome_produto,
                        'Cor/Variação': descricao_cor if descricao_cor else "",
                        'Qtde em Estoque': qtde_estoque,
                        'Qtde Reservada': reservada,
                        'Qtde Disponível': qtde_estoque - reservada,
                        'Página': pagina_num
                    })
                    if reservada > 0 and pagina_num not in paginas_com_reserva:
                        paginas_com_reserva.append(pagina_num)
                cod_produto = None
                nome_produto = None
                descricao_cor = None
                qtde_estoque = None
                qtde_reservada = None

            for pagina in pdf.pages:
                pagina_num += 1
                texto = pagina.extract_text()

                if texto:
                    linhas = texto.split('\n')

                    for i, linha in enumerate(linhas):
                        linha_limpa = linha.strip()

                        cabecalho = extrair_cabecalho_produto(linha_limpa)

                        if cabecalho:
                            flush_produto()
                            cod_produto, nome_produto = cabecalho

                        elif cod_produto and re.match(r'^\d{3}\s*-\s*.+', linha_limpa):
                            descricao_cor = linha_limpa

                        elif linha_limpa.startswith('Qtde em Estoque'):
                            partes = linha_limpa.split()
                            if len(partes) >= 4:
                                try:
                                    qtde_estoque = float(partes[-1].replace(',', '.'))
                                except:
                                    pass

                        elif linha_limpa.startswith('Qtde Total em Estoque'):
                            partes = linha_limpa.split()
                            try:
                                total_geral_estoque = float(partes[-1].replace(',', '.'))
                            except:
                                pass

                        elif linha_limpa.startswith('Qtde Total Reservada'):
                            partes = linha_limpa.split()
                            try:
                                total_geral_reservada = float(partes[-1].replace(',', '.'))
                            except:
                                pass

                        elif linha_limpa.startswith('Qtde Reservada'):
                            partes = linha_limpa.split()
                            if len(partes) >= 3:
                                try:
                                    qtde_reservada = float(partes[-1].replace(',', '.'))
                                except:
                                    qtde_reservada = 0

            flush_produto()

            for pag in sorted(set(paginas_com_reserva)):
                prods_pag = [p for p in produtos if p['Página'] == pag and p['Qtde Reservada'] > 0]
                if prods_pag:
                    log.write(f"PÁGINA {pag} - PRODUTOS COM QUANTIDADE RESERVADA:\n")
                    log.write("-"*80 + "\n")
                    for prod in prods_pag:
                        log.write(f"  Código: {prod['Código']}\n")
                        log.write(f"  Nome: {prod['Nome do Produto']}\n")
                        log.write(f"  Cor/Variação: {prod['Cor/Variação']}\n")
                        log.write(f"  Estoque: {prod['Qtde em Estoque']} | Reservada: {prod['Qtde Reservada']} | Disponível: {prod['Qtde Disponível']}\n")
                        log.write("\n")
        
        log.write("="*80 + "\n")
        log.write("RESUMO\n")
        log.write("="*80 + "\n")
        log.write(f"Total de produtos encontrados: {len(produtos)}\n")
        
        produtos_com_reserva = [p for p in produtos if p['Qtde Reservada'] > 0]
        log.write(f"Produtos com quantidade reservada: {len(produtos_com_reserva)}\n")
        
        if paginas_com_reserva:
            log.write(f"Páginas com produtos reservados: {', '.join(map(str, sorted(paginas_com_reserva)))}\n")
        
        total_estoque = sum(p['Qtde em Estoque'] for p in produtos)
        total_reservado = sum(p['Qtde Reservada'] for p in produtos)
        
        log.write(f"Total de unidades em estoque: {total_estoque}\n")
        log.write(f"Total de unidades reservadas: {total_reservado}\n")
        if total_geral_estoque is not None:
            log.write(f"Total geral em estoque (impresso no PDF): {total_geral_estoque}\n")
        if total_geral_reservada is not None:
            log.write(f"Total geral reservado (impresso no PDF): {total_geral_reservada}\n")
        log.write("="*80 + "\n")

    return produtos, total_geral_estoque, total_geral_reservada

def salvar_resultados(produtos, arquivo_excel):
    df = pd.DataFrame(produtos)

    df.to_excel(arquivo_excel, index=False)
    print(f"Dados salvos em Excel: {arquivo_excel}")

    return df

def exibir_resumo(df):
    print("\n" + "="*80)
    print("RESUMO DO INVENTÁRIO")
    print("="*80)
    print(f"Total de produtos: {len(df)}")
    print(f"Total em Estoque: {df['Qtde em Estoque'].sum()}")
    print(f"Total Reservado: {df['Qtde Reservada'].sum()}")
    print(f"Total Disponível: {df['Qtde Disponível'].sum()}")
    
    produtos_com_reserva = df[df['Qtde Reservada'] > 0]
    print(f"\nProdutos com quantidade reservada: {len(produtos_com_reserva)}")
    
    if len(produtos_com_reserva) > 0:
        paginas = sorted(produtos_com_reserva['Página'].unique())
        print(f"Páginas com produtos reservados: {', '.join(map(str, paginas))}")
    
    print("="*80)
    
    print("\nPrimeiros 10 produtos:")
    print(df.head(10).to_string(index=False))
    
    print("\nÚltimos 5 produtos:")
    print(df.tail(5).to_string(index=False))
    
    if len(produtos_com_reserva) > 0:
        print("\n" + "="*80)
        print("PRODUTOS COM QUANTIDADE RESERVADA:")
        print("="*80)
        print(produtos_com_reserva[['Código', 'Nome do Produto', 'Qtde em Estoque', 'Qtde Reservada', 'Página']].to_string(index=False))

if __name__ == "__main__":
    caminho_pdf = r"C:\Users\ricardo\Documents\GitHub\Inventario\Inventario\61 MENDES.pdf"
    pasta_pdf = os.path.dirname(caminho_pdf)
    arquivo_log = os.path.join(pasta_pdf, "61_MENDES_reservas.log")
    arquivo_excel = os.path.join(pasta_pdf, "61_MENDES_inventario.xlsx")

    print("Iniciando extração de dados do PDF...")
    print(f"Arquivo: {caminho_pdf}\n")

    produtos, total_geral_estoque, total_geral_reservada = extrair_produtos_inventario(caminho_pdf, arquivo_log)

    if produtos:
        print(f"\nArquivo de log criado: {arquivo_log}")

        df = salvar_resultados(produtos, arquivo_excel)

        exibir_resumo(df)

        if total_geral_estoque is not None and total_geral_reservada is not None:
            print(f"\n{'='*80}")
            print("VERIFICAÇÃO (comparado com o total impresso no PDF):")
            print(f"Total esperado em estoque (PDF): {total_geral_estoque}")
            print(f"Total calculado em estoque: {df['Qtde em Estoque'].sum()}")
            print(f"Diferença: {df['Qtde em Estoque'].sum() - total_geral_estoque}")
            print(f"\nTotal esperado reservado (PDF): {total_geral_reservada}")
            print(f"Total calculado reservado: {df['Qtde Reservada'].sum()}")
            print(f"Diferença: {df['Qtde Reservada'].sum() - total_geral_reservada}")
            print(f"{'='*80}")
    else:
        print("Nenhum produto foi encontrado no PDF.")