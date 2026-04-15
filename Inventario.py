import pdfplumber
import pandas as pd
import re
from datetime import datetime

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
            
            codigo_produto = None
            descricao_cor = None
            qtde_estoque = None
            qtde_reservada = None
            pagina_num = 0
            
            for pagina in pdf.pages:
                pagina_num += 1
                texto = pagina.extract_text()
                
                if texto:
                    linhas = texto.split('\n')
                    
                    for i, linha in enumerate(linhas):
                        linha_limpa = linha.strip()
                        
                        if re.match(r'^\d{4,}\s*-\s*.+', linha_limpa):
                            if codigo_produto and qtde_estoque is not None:
                                if qtde_reservada is None:
                                    qtde_reservada = 0
                                
                                match_codigo = re.match(r'^(\d+)\s*-\s*(.+)', codigo_produto)
                                if match_codigo:
                                    cod = match_codigo.group(1)
                                    nome = match_codigo.group(2).strip()
                                else:
                                    cod = codigo_produto
                                    nome = ""
                                
                                produto = {
                                    'Código': cod,
                                    'Nome do Produto': nome,
                                    'Cor/Variação': descricao_cor if descricao_cor else "",
                                    'Qtde em Estoque': qtde_estoque,
                                    'Qtde Reservada': qtde_reservada,
                                    'Qtde Disponível': qtde_estoque - qtde_reservada,
                                    'Página': pagina_num
                                }
                                produtos.append(produto)
                                
                                if qtde_reservada > 0:
                                    if pagina_num not in paginas_com_reserva:
                                        paginas_com_reserva.append(pagina_num)
                            
                            codigo_produto = linha_limpa
                            descricao_cor = None
                            qtde_estoque = None
                            qtde_reservada = None
                        
                        elif codigo_produto and re.match(r'^\d{3}\s*-\s*.+', linha_limpa):
                            descricao_cor = linha_limpa
                        
                        elif linha_limpa.startswith('Qtde em Estoque'):
                            partes = linha_limpa.split()
                            if len(partes) >= 4:
                                try:
                                    qtde_estoque = int(partes[-1])
                                except:
                                    pass
                        
                        elif linha_limpa.startswith('Qtde Reservada'):
                            partes = linha_limpa.split()
                            if len(partes) >= 3:
                                try:
                                    qtde_reservada = int(partes[-1])
                                except:
                                    qtde_reservada = 0
                    
                    if codigo_produto and qtde_estoque is not None:
                        if qtde_reservada is None:
                            qtde_reservada = 0
                        
                        match_codigo = re.match(r'^(\d+)\s*-\s*(.+)', codigo_produto)
                        if match_codigo:
                            cod = match_codigo.group(1)
                            nome = match_codigo.group(2).strip()
                        else:
                            cod = codigo_produto
                            nome = ""
                        
                        produto = {
                            'Código': cod,
                            'Nome do Produto': nome,
                            'Cor/Variação': descricao_cor if descricao_cor else "",
                            'Qtde em Estoque': qtde_estoque,
                            'Qtde Reservada': qtde_reservada,
                            'Qtde Disponível': qtde_estoque - qtde_reservada,
                            'Página': pagina_num
                        }
                        produtos.append(produto)
                        
                        if qtde_reservada > 0:
                            if pagina_num not in paginas_com_reserva:
                                paginas_com_reserva.append(pagina_num)
                        
                        codigo_produto = None
                        descricao_cor = None
                        qtde_estoque = None
                        qtde_reservada = None
            
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
        log.write("="*80 + "\n")
    
    return produtos

def salvar_resultados(produtos, arquivo_excel, arquivo_csv):
    df = pd.DataFrame(produtos)
    
    df.to_excel(arquivo_excel, index=False)
    print(f"Dados salvos em Excel: {arquivo_excel}")
    
    df.to_csv(arquivo_csv, index=False, encoding='utf-8-sig')
    print(f"Dados salvos em CSV: {arquivo_csv}")
    
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
    caminho_pdf = r"C:\Users\ricardo\Documents\Inventario\61 MENDES.pdf"
    arquivo_log = r"C:\Users\ricardo\Documents\Inventario\61_MENDES_reservas.log"
    
    print("Iniciando extração de dados do PDF...")
    print(f"Arquivo: {caminho_pdf}\n")
    
    produtos = extrair_produtos_inventario(caminho_pdf, arquivo_log)
    
    if produtos:
        print(f"\nArquivo de log criado: {arquivo_log}")
        
        df = salvar_resultados(
            produtos,
            r"C:\Users\ricardo\Documents\Inventario\61_MENDES_inventario.xlsx",
            r"C:\Users\ricardo\Documents\Inventario\61_MENDES_inventario.csv"
        )
        
        exibir_resumo(df)
        
        print(f"\n{'='*80}")
        print("VERIFICAÇÃO:")
        print(f"Total esperado em estoque (PDF): 90578")
        print(f"Total calculado em estoque: {df['Qtde em Estoque'].sum()}")
        print(f"Diferença: {df['Qtde em Estoque'].sum() - 90578}")
        print(f"\nTotal esperado reservado (PDF): 42")
        print(f"Total calculado reservado: {df['Qtde Reservada'].sum()}")
        print(f"Diferença: {df['Qtde Reservada'].sum() - 42}")
        print(f"{'='*80}")
    else:
        print("Nenhum produto foi encontrado no PDF.")