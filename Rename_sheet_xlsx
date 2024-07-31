import os
import openpyxl

# Função para renomear a planilha em um arquivo Excel
def rename_sheet(file_path, old_names, new_name):
    # Verifica se o arquivo existe
    if os.path.exists(file_path):
        # Carrega o arquivo Excel
        wb = openpyxl.load_workbook(file_path)
        
        # Verifica se a planilha com um dos nomes antigos existe
        found = False
        for old_name in old_names:
            if old_name in wb.sheetnames:
                # Renomeia a planilha
                ws = wb[old_name]
                ws.title = new_name
                
                # Salva as alterações no arquivo Excel
                wb.save(file_path)
                print(f'A planilha "{old_name}" no arquivo "{file_path}" foi renomeada para "{new_name}" com sucesso!')
                found = True
                break
        
        if not found:
            print(f'Nenhuma das planilhas {" ou ".join([f"\"{name}\"" for name in old_names])} foi encontrada no arquivo "{file_path}".')
    else:
        print(f'O arquivo "{file_path}" não existe.')

# Pasta onde estão os arquivos
pasta = r'C:\EngDados\dados\Planilhas_Fixo'

# Nomes antigos e novo da planilha
old_sheet_names = ['TELEFONIA _ FIXA', 'TELEFONIA FIXA', 'TELEFONIA _FIXA', 'TELEFONIA - FIXA']
new_sheet_name = 'TELEFONIA_FIXA'

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.xlsx'):  # Verifica se é um arquivo Excel
        caminho_arquivo = os.path.join(pasta, arquivo)
        rename_sheet(caminho_arquivo, old_sheet_names, new_sheet_name)
