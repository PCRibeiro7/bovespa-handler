import numpy as np
import xlsxwriter

wb = xlsxwriter.Workbook('planilha.xlsx')
ws = wb.add_worksheet('A Test Sheet')
arq=open('historico.txt','r')
current_line= arq.readline()
indices_header=[[0,2],[2,15],[15,23],[23,31],[31,245]]
indices_cotacoes=[[0,2],[2,10],[10,12],[12,24],[24,27],[27,39],[39,49],[49,52],[52,56],[56,69],[69,82],[82,95],[95,108],[108,121],[121,134],[134,147],[147,152],[152,170],[170,188],[188,201],[201,202],[202,210],[210,217],[217,230],[230,242],[242,245]]
indices_trailer=[[0,2],[2,15],[15,23],[23,31],[31,42],[42,245]]
linha_index=0
numero_max=10000

while current_line!='':
        add_line=[]
        coluna_index=0
        if current_line[0:2]=='00':
                for indice_current in indices_header:
                        ws.write(linha_index, coluna_index,current_line[indice_current[0]:indice_current[1]])
                        coluna_index+=1
        if current_line[0:2]=='01':
                for indice_current in indices_cotacoes:
                        ws.write(linha_index, coluna_index,current_line[indice_current[0]:indice_current[1]])
                        coluna_index+=1
        if current_line[0:2]=='99':
                for indice_current in indices_trailer:
                        ws.write(linha_index, coluna_index,current_line[indice_current[0]:indice_current[1]])
                        coluna_index+=1                
        current_line=arq.readline()
        if linha_index==numero_max:
            print ('Indexando a linha:',linha_index)
            break
        linha_index+=1
wb.close()
        



