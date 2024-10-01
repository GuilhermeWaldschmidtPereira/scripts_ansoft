"""
Este script processa arquivos de texto contendo coordenadas de pontos e gera scripts no formato .vbs para serem utilizados no Ansoft Designer.
O script realiza as seguintes etapas:
1. Define os nomes dos arquivos que serão processados.
2. Define os caminhos de entrada e saída para os arquivos.
3. Lê o arquivo de entrada para contar a quantidade de linhas (número de pontos).
4. Lê o arquivo de entrada novamente para extrair as coordenadas x e y.
5. Gera um script no formato .vbs com as coordenadas extraídas.
6. Escreve o script gerado em um arquivo de saída.
Variáveis:
- name_file: Nome do arquivo a ser processado (sem extensão).
- input_file_path: Caminho do arquivo de entrada contendo as coordenadas dos pontos.
- output_file_path: Caminho do arquivo de saída onde o script .vbs será salvo.
- line_points: Lista para armazenar as coordenadas dos pontos.
- num_linhas: Número de linhas (pontos) no arquivo de entrada.
- line_data: String contendo o conteúdo do script .vbs gerado.
Dependências:
- O script assume que os arquivos de entrada estão no formato de texto com duas colunas de valores numéricos separados por espaço.
- O script gerado é específico para ser utilizado no Ansoft Designer Version 3.5.0.
Uso:
- Atualize a lista `name_file` com os nomes dos arquivos a serem processados.
- Certifique-se de que os caminhos de entrada e saída estão corretos.
- Execute o script para gerar os arquivos .vbs correspondentes.

@Author: Guilherme Waldschmidt Pereira
@Date: 2024-09

"""


# Colocar os nomes dos arquivos que serão processados
for name_file in [
    'espiral_fib1'
]:
    # Aqui o input e o output são definidos (colocar o path correto)
    input_file_path = f'espiral_unitaria/nova_versao/{name_file}.txt'
    output_file_path = f'espiral_unitaria/nova_versao/{name_file}.vbs'
    
    # Pega a quantidade de linhas (número de pontos)
    with open(f'.\{input_file_path}', 'r') as file:
        line_points = []
        num_linhas = sum(1 for line in file)

    # Lê o arquivo
    with open(f'.\{input_file_path}', 'r') as file:
        line_data = """
        ' ----------------------------------------------
        ' Script Recorded by Ansoft Designer Version 3.5.0
        ' 12:58 AM  Sep 27, 2024
        ' ----------------------------------------------
        Dim oAnsoftApp
        Dim oDesktop
        Dim oProject
        Dim oDesign
        Dim oEditor
        Dim oModule
        Set oAnsoftApp = CreateObject("AnsoftDesigner.DesignerScript")
        Set oDesktop = oAnsoftApp.GetAppDesktop()
        oDesktop.RestoreWindow
        Set oProject = oDesktop.SetActiveProject("teste")
        Set oDesign = oProject.SetActiveDesign("PlanarEM1")
        Set oEditor = oDesign.SetActiveEditor("Layout")


        oEditor.CreateLine Array("NAME:Contents", "lineGeometry:=", Array("Name:=", "line_20", "LayerName:=",  _\n  "Trace", "lw:=", "30mil", "endstyle:=", 0, "joinstyle:=", 0, "n:=", {}, "U:=",  _\n  "mm", """.format(num_linhas)
        
        # Grava as coordenadas x e y do arquivo no padrão aceito pelo script
        for line in file:
            col1, col2 = line.split()  # Aqui é preciso alterar se houver outro valor separando as colunas
            line_points.append(f'"x:=", {float(col1)}, "y:=", {float(col2)}')

    line_data += ', '.join(line_points) + "))"

    # Escreve o arquivo de saída no formato .vbs
    with open(output_file_path, 'w') as output_file:
        output_file.write(line_data)
