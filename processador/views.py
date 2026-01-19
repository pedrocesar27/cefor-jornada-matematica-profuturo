from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
import pandas as pd
import os
from pathlib import Path


def index(request):
    """Página inicial com formulário de upload"""
    return render(request, 'processador/index.html')


def processar_arquivo(request):
    if request.method == 'POST' and request.FILES.getlist('arquivo'):
        arquivos = request.FILES.getlist('arquivo')
        fs = FileSystemStorage()

        resultados_gerais = []

        try:
            for arquivo in arquivos:
                # Salva cada arquivo temporariamente
                filename = fs.save(arquivo.name, arquivo)
                arquivo_path = fs.path(filename)

                abas = pd.read_excel(arquivo_path, sheet_name=None)

                for nome_aba, df in abas.items():
                    # Normalizações
                    df = df.rename(columns={'DRE ': 'DRE', 'CH total ': 'CH total'})

                    if df["CH cursada"].dtype == object:
                        df["CH cursada"] = pd.to_timedelta(
                            df["CH cursada"] + ":00"
                        ).dt.total_seconds() / 3600

                    # Preenche células mescladas
                    df['DRE'] = df['DRE'].ffill()
                    df['Nome da escola'] = df['Nome da escola'].ffill()

                    # Agrupa dentro do próprio arquivo
                    soma_por_cpf = df.groupby(
                        ["CPF", "DRE", "Nome da escola", "Cursista"],
                        as_index=False
                    )["CH cursada"].sum()

                    resultados_gerais.append(soma_por_cpf)

                # Remove o arquivo temporário depois de processar
                fs.delete(filename)

            # 🔥 AQUI acontece o merge real entre todos os arquivos
            resultado_final = pd.concat(resultados_gerais, ignore_index=True)

            resultado_final = resultado_final.groupby(
                ["CPF", "DRE", "Nome da escola", "Cursista"],
                as_index=False
            )["CH cursada"].sum()

            resultado_final = resultado_final.sort_values(["CPF", "DRE"])

            # Salva o Excel final unificado
            output_path = fs.path('horas_por_cpf.xlsx')
            resultado_final.to_excel(output_path, index=False)

            dados = resultado_final.to_dict('records')
            total_cpfs = resultado_final["CPF"].nunique()
            total_horas = resultado_final["CH cursada"].sum()

            context = {
                'sucesso': True,
                'dados': dados,
                'total_cpfs': total_cpfs,
                'total_horas': round(total_horas, 2)
            }

            return render(request, 'processador/resultado.html', context)

        except Exception as e:
            context = {'erro': str(e)}
            return render(request, 'processador/index.html', context)

    return render(request, 'processador/index.html')



def download_resultado(request):
    """Permite baixar o arquivo processado"""
    fs = FileSystemStorage()
    output_path = fs.path('horas_por_cpf.xlsx')
    
    if os.path.exists(output_path):
        return FileResponse(
            open(output_path, 'rb'),
            as_attachment=True,
            filename='horas_por_cpf.xlsx'
        )
    else:
        return HttpResponse('Arquivo não encontrado', status=404)
