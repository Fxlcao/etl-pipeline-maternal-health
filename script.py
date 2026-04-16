import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def ler_abas(caminho):
    print(f"Lendo arquivo de entrada: {caminho}")
    # Carrega todas as planilhas do Excel para a memória em um dicionário de DataFrames
    xl = pd.ExcelFile(caminho)
    abas = {}
    for nome in xl.sheet_names:
        abas[nome] = pd.read_excel(caminho, sheet_name=nome, header=0)
    return abas

def val(row, col, padrao=None):
    # Função auxiliar para extração segura de valores, evitando falhas em colunas inexistentes ou vazias
    if row is None: return padrao
    try:
        v = row.get(col, padrao)
        return padrao if pd.isna(v) else v
    except:
        return padrao

def montar_juncao(abas):
    epds      = abas.get('Cadastro EPDS')
    enfe_full = abas.get('Instrumento de Consulta de Enfe')
    estrat    = abas.get('Estratificação de Risco Gestaci')
    estrat2   = abas.get('Estratificação de Risco Ges (2)')
    cad       = abas.get('Cadastro Pacientes')

    if epds is None:
        messagebox.showerror("Erro", "Aba 'Cadastro EPDS' não encontrada!")
        return pd.DataFrame()

    registros = []
    
    # Itera sobre a base principal (EPDS), preservando duplicidades válidas (ex: partos múltiplos)
    for index, r_epds in epds.iterrows():
        nome = r_epds['Paciente.Nome Social']
        if pd.isna(nome): continue

        # Cruzamento de dados (Left Join manual) utilizando o Nome Social como chave de ligação
        r_enfe = enfe_full[enfe_full['Paciente.Nome Social'] == nome].iloc[0] if enfe_full is not None and not enfe_full[enfe_full['Paciente.Nome Social'] == nome].empty else None
        r_cad  = cad[cad['Nome Social'] == nome].iloc[0] if cad is not None and not cad[cad['Nome Social'] == nome].empty else None
        
        r_estrat = None
        if estrat is not None and not estrat[estrat['Paciente.Nome Social'] == nome].empty:
            r_estrat = estrat[estrat['Paciente.Nome Social'] == nome].iloc[-1]
        elif estrat2 is not None and not estrat2[estrat2['Paciente.Nome Social'] == nome].empty:
            r_estrat = estrat2[estrat2['Paciente.Nome Social'] == nome].iloc[-1]

        # Mapeamento do esquema de dados consolidado
        registro = {
            'NAME': nome,
            'AGE': val(r_cad, 'Idade'),
            'RACE': val(r_cad, 'Raça/Cor'),
            'NACIONALITY': val(r_cad, 'Nacionalidade'),
            'COUNTRY': val(r_cad, 'País de Nascim'),
            'CITY': val(r_cad, 'Município'),
            'ESCOLARITY': val(r_enfe, 'Escolaridade'),
            'HEIGHT': val(r_enfe, 'Altura (cm)'),
            'WEIGHT': val(r_enfe, 'Peso Pré-Gesta'),
            'PRE-PREGNANCY WEIGHT': val(r_enfe, 'Peso Pré-Gesta'),
            'IBM': val(r_enfe, 'calculo imc'),
            'GESTATIONAL AGE - weeks': val(r_enfe, 'Idade Gestacional'),
            'GESTATIONAL AGE - days': val(r_enfe, 'Idade Gestacional _x'),
            'PREVIOUS PREGNANCY': val(r_enfe, 'Gestações Pr_x'),
            'Cesaria Anterior': val(r_enfe, 'Cesáreas Prévi'),
            'Abortos Previos': val(r_enfe, 'Abortos Prévios'),
            'Gestação desejada': val(r_enfe, 'Gestação desej'),
            'Prematuridade Gestaçao Anterior': val(r_enfe, 'Prematuridade na ges'),
            'Partos Normais Previos': val(r_enfe, 'Partos normais pr_x0'),
            'Nascidos Vivos': val(r_enfe, 'Nascidos vivos'),
            'Natimortos': val(r_enfe, 'Natimortos'),
            'Isoimunização ': val(r_enfe, 'Isoimunização '),
            'Idade Filho mais velho': val(r_enfe, 'Idade filho mais_x00'),
            'Metodo Contracepitivo Anteriormente': val(r_enfe, 'Utilizou método_x002'),
            'Se Sim Qual?': val(r_enfe, 'Se sim, quais_0'),
            'Idade Inicio Atv. Sexual': val(r_enfe, 'Idade de iníci'),
            'Intercorrencia': val(r_enfe, 'Intercorrências como'),
            'Medida Proteçao': val(r_enfe, 'Medidas de prote_x00'),
            'Ano do ultimo preventivo': val(r_enfe, 'Ano do último_'),
            'Realizou Testes Rapidos': val(r_enfe, 'Realizou testes r_x0'),
            'Se Sim, Data': val(r_enfe, 'Se sim, data_x'),
            'Data da ultima': val(r_enfe, 'Data da última'),
            'Antecedentes Gine e obstreticos': val(r_enfe, 'Antecedentes ginecol_x00f30'),
            'Historico/antecedente': val(r_enfe, 'Histórico/Antecedent'),
            'Historico Familiar': val(r_enfe, 'Histórico familiar_x'),
            'Se sim Qual': val(r_enfe, 'Se sim, qual_x'),
            'Historico Cirurgia': val(r_enfe, 'Histórico de c'),
            'Saúde bucal _x': val(r_enfe, 'Saúde bucal _x'),
            'Etilismo': val(r_enfe, 'Etilismo'),
            'Tabagismo': val(r_enfe, 'Tabagismo'),
            'Exposição  Fumaca do cigarro': val(r_enfe, 'Exposição _x00'),
            'Outras drogas': val(r_enfe, 'Outras drogas'),
            'Se sim, quais_': val(r_enfe, 'Se sim, quais_'),
            'Ativ. Fisica': val(r_enfe, 'Pratica atividade f_'),
            'se Sim Qual?2': val(r_enfe, 'Se sim, de_x00'),
            'Há histórico familiar de depressão pós-parto?': val(r_epds, 'Há histórico_x'),
            'Qual a patologia?': val(r_epds, 'Há histórico_x0'),
            'Realiza tratamento?': val(r_epds, 'Realiza tratamento?'),
            'Qual o tratamento': val(r_epds, 'Qual o tratamento_x0'),
            'Faz uso de drogas ilicitas': val(r_epds, 'Faz uso de dro'),
            'Faz uso de drogas licitas': val(r_epds, 'Faz uso de dro0'),
            'Há histórico de internação psiquiátrica?': val(r_epds, 'Há histórico_x1'),
            'Eu tenho sido capaz de rir e achar graça das coisas?': val(r_epds, 'OData_1. Eu te'),
            '\xa0Eu tenho pensado no futuro com alegria': val(r_epds, 'Sim, como de_x'),
            'Eu tenho me culpado sem razão quando as coisas dão errado': val(r_epds, 'Não, de'),
            'Eu tenho ficado ansiosa ou preocupada sem uma boa razão': val(r_epds, 'Sim, muito seg'),
            ' Eu tenho me sentido assustada ou em pânico sem um bom motivo': val(r_epds, 'OData_5. Eu te'),
            'Eu tenho me sentido sobrecarregada pelas tarefas e acontecimentos do meu dia-a-dia': val(r_epds, 'OData_6. Eu te'),
            'Eu tenho me sentido tão infeliz que eu tenho tido dificuldade de dormir': val(r_epds, 'OData_7. Eu te'),
            'Eu tenho me sentido triste ou muito mal': val(r_epds, 'OData_8. Eu te'),
            'Eu tenho me sentido tão triste que tenho chorado': val(r_epds, 'OData_9. Eu te'),
            'Eu tenho pensado em fazer alguma coisa contra mim mesma': val(r_epds, 'OData_10. Eu t'),
            'Resultado': val(r_epds, 'Resultado'),
            'Foi Reestratificada?': val(r_estrat, 'Foi Reestratificada?'),
            'Pontuação da_x': val(r_estrat, 'Pontuação da_x'),
            'Faixa de Idade': val(r_estrat, 'Faixa de Idade'),
            'Mulher de Raça': val(r_estrat, 'Mulher de Raça'),
            'Baixa Escolaridade': val(r_estrat, 'Baixa Escolaridade'),
            'Tabagista Ativa': val(r_estrat, 'Tabagista Ativa'),
            'Indícios de oc': val(r_estrat, 'Indícios de oc'),
            'Gestante em situa_x0': val(r_estrat, 'Gestante em situa_x0'),
            'Avaliação Nutr': val(r_estrat, 'Avaliação Nutr'),
            'OData_2 Abortos cons': val(r_estrat, 'OData_2 Abortos cons'),
            'OData_3 Abortos n_x0': val(r_estrat, 'OData_3 Abortos n_x0'),
            'Prematuridade na ges': val(r_estrat, 'Prematuridade na ges'),
            'Mais de um par': val(r_estrat, 'Mais de um par'),
            'Restrição de_x': val(r_estrat, 'Restrição de_x'),
            'Natimorto sem causa_': val(r_estrat, 'Natimorto sem causa_'),
            'Incompetência istmo_': val(r_estrat, 'Incompetência istmo_'),
            'Isoimunização 2': val(r_estrat, 'Isoimunização '),
            'Pré-Eclâmpsia_': val(r_estrat, 'Pré-Eclâmpsia_'),
            'Psicose puerperal na': val(r_estrat, 'Psicose puerperal na'),
            'Transplante': val(r_estrat, 'Transplante'),
            'Cirurgia bariátrica_': val(r_estrat, 'Cirurgia bariátrica_'),
            'Acretismo placentári': val(r_estrat, 'Acretismo placentári'),
            'Doença hipertensiva_': val(r_estrat, 'Doença hipertensiva_'),
            'Diabetes gestacional_x0020': val(r_estrat, 'Diabetes gestacional_x0020'),
            'Infecção urin_': val(r_estrat, 'Infecção urin_'),
            'Cálculo renal ': val(r_estrat, 'Cálculo renal '),
            'Restrição de_x0': val(r_estrat, 'Restrição de_x0'),
            'Feto acima do ': val(r_estrat, 'Feto acima do '),
            'Polidrâmno/Oligodr_x': val(r_estrat, 'Polidrâmno/Oligodr_x'),
            'Colo curto em ': val(r_estrat, 'Colo curto em '),
            'Suspeita de acretism': val(r_estrat, 'Suspeita de acretism'),
            'Placenta prévia_x002': val(r_estrat, 'Placenta prévia_x002'),
            'Hepatopatias (Ex_x00': val(r_estrat, 'Hepatopatias (Ex_x00'),
            'Anemia grave ou_x002': val(r_estrat, 'Anemia grave ou_x002'),
            'Isoimunização 0': val(r_estrat, 'Isoimunização 0'),
            'Câncer materno_x0020': val(r_estrat, 'Câncer materno_x0020'),
            'Neoplasias ginecológ': val(r_estrat, 'Neoplasias ginecológ'),
            'Alta suspeita cl_x00': val(r_estrat, 'Alta suspeita cl_x00'),
            'Lesão de alto_': val(r_estrat, 'Lesão de alto_'),
            'Suspeita de malforma': val(r_estrat, 'Suspeita de malforma'),
            'Gemelaridade': val(r_estrat, 'Gemelaridade'),
            'Sífilis (Terci': val(r_estrat, 'Sífilis (Terci'),
            'Condiloma acuminado ': val(r_estrat, 'Condiloma acuminado '),
            'Hepatites agudas com': val(r_estrat, 'Hepatites agudas com'),
            'Hanseníase com_x0020': val(r_estrat, 'Hanseníase com_x0020'),
            'AIDS/HIV com d': val(r_estrat, 'AIDS/HIV com d'),
            'Tuberculose': val(r_estrat, 'Tuberculose'),
            'Toxoplasmose ou rub_': val(r_estrat, 'Toxoplasmose ou rub_'),
            'Dependência e/': val(r_estrat, 'Dependência e/'),
            'Endocrinopatias descompens': val(r_estrat, 'Endocrinopatias descompens'),
            'Suspeita ou confirma': val(r_estrat, 'Suspeita ou confirma'),
            'Suspeita ou confirma0': val(r_estrat, 'Suspeita ou confirma0'),
            'Hipertensão arterial': val(r_estrat, 'Hipertensão arterial'),
            'Diabetes mellitus 1_': val(r_estrat, 'Diabetes mellitus 1_'),
            'Tireoidopatias (hipe': val(r_estrat, 'Tireoidopatias (hipe'),
            'Doença psiquiá': val(r_estrat, 'Doença psiquiá'),
            'Doenças hematol_x00f': val(r_estrat, 'Doenças hematol_x00f'),
            'Cardiopatias com rep': val(r_estrat, 'Cardiopatias com rep'),
            'Pneumopatias graves ': val(r_estrat, 'Pneumopatias graves '),
            'Doenças auto-i': val(r_estrat, 'Doenças auto-i'),
            'Uso de medicamentos_': val(r_estrat, 'Uso de medicamentos_'),
            'Doença renal g': val(r_estrat, 'Doença renal g'),
            'Hemopatias e anemia_': val(r_estrat, 'Hemopatias e anemia_'),
            'Hepatopatias crônica': val(r_estrat, 'Hepatopatias crônica'),
            'Pontuacao Estratificacao': val(r_estrat, 'Pontuação da_x0'),
        }
        registros.append(registro)

    df = pd.DataFrame(registros)

    # --- Transformações de Limpeza e Padronização de Dados ---

    # Extrai a parte inteira da pontuação por meio de split na string
    if 'Pontuacao Estratificacao' in df.columns:
        df['Pontuacao Estratificacao'] = df['Pontuacao Estratificacao'].apply(
            lambda x: str(x).split('.')[0] if pd.notna(x) and str(x).strip() != '' else x
        )

    # Imputação condicional de valores nulos (Null Handling) baseada em inferência de tipo
    for col in df.columns:
        col_temp = pd.to_numeric(df[col], errors='ignore')
        
        if pd.api.types.is_numeric_dtype(col_temp):
            df[col] = df[col].fillna("Null")
        else:
            # Normaliza strings vazias ou compostas apenas por espaços para o tipo nulo padrão do Pandas
            df[col] = df[col].replace(r'^\s*$', pd.NA, regex=True)
            df[col] = df[col].fillna("-")

    return df

def salvar_csv(df, caminho_saida):
    try:
        # Exportação do dataset em formato flat-file com codificação suportada nativamente por ferramentas de BI e Excel
        df.to_csv(caminho_saida, index=False, sep=';', encoding='utf-8-sig')
        messagebox.showinfo("Sucesso", f"Arquivo CSV gerado com sucesso!\nSalvo em: {caminho_saida}")
    except PermissionError:
        messagebox.showerror("Erro", "O arquivo de destino está bloqueado/aberto. Feche-o e tente novamente.")

def selecionar_arquivos():
    root = tk.Tk()
    root.withdraw()
    
    # Interface de seleção interativa para ingestão e extração
    c_entrada = filedialog.askopenfilename(title="Selecionar Origem de Dados (Excel)", filetypes=[("Excel", "*.xlsx")])
    if not c_entrada: return
    
    c_saida = filedialog.asksaveasfilename(
        title="Definir Destino do Dataset (CSV)", 
        defaultextension=".csv", 
        filetypes=[("CSV", "*.csv")],
        initialfile="Dataset_Consolidado_AGAR.csv"
    )
    if not c_saida: return

    try:
        abas = ler_abas(c_entrada)
        df_final = montar_juncao(abas)
        if not df_final.empty:
            salvar_csv(df_final, c_saida)
    except Exception as e:
        messagebox.showerror("Erro de Execução", str(e))
    root.destroy()

if __name__ == '__main__':
    selecionar_arquivos()
