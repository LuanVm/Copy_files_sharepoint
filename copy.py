import os
import glob
import shutil

# Caminho do diretório principal
diretorio_principal = r'' #Origem_dados, onedrive ou sharepoint devem estar linkados na máquina e não na web

# Caminho para o diretório de destino
diretorio_destino = r'' #Destino_armazenar

# Percorre cada subdiretório dentro do diretório principal
for pasta_cliente in os.listdir(diretorio_principal):
    caminho_pasta_cliente = os.path.join(diretorio_principal, pasta_cliente)
    
    # Verifica se é um diretório
    if os.path.isdir(caminho_pasta_cliente):
        # Caminho para a subpasta 'Estrutura'
        caminho_Estrutura = os.path.join(caminho_pasta_cliente, 'Estrutura')
        
        # Verifica se a subpasta 'Estrutura' existe
        if os.path.exists(caminho_Estrutura):
            # Lista todos os arquivos Excel na subpasta 'Estrutura'
            arquivos_excel = glob.glob(os.path.join(caminho_Estrutura, '*.xlsx'))
            
            # Copia os arquivos Excel encontrados para o diretório de destino
            for arquivo_excel in arquivos_excel:
                shutil.copy(arquivo_excel, diretorio_destino)

print("Arquivos copiados para", diretorio_destino)
