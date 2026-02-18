import pandas as pd
import unicodedata
from Emailer import Emailer
import streamlit as st

def remover_acentos(text):
    # NFD separa o acento da letra (ex: ç -> c + cedilha)
    text = text.replace("'","")
    normalized = unicodedata.normalize('NFD', text)
    # Ignora os caracteres que não são ASCII (acentos)
    return normalized.encode('ascii', 'ignore').decode('utf-8')

def acha_mes(mes):
    if mes =='jan.':
        return '01'
    if mes =='fev.':
        return '02'
    if mes =='mar.':
        return '03'
    if mes =='abr.':
        return '04'
    if mes =='mai.':
        return '05'
    if mes =='jun.':
        return '06'
    if mes =='jul.':
        return '07'
    if mes =='ago.':
        return '08'
    if mes =='set.':
        return '09'
    if mes =='out.':
        return '10'
    if mes =='nov.':
        return '11'
    if mes =='dez.':
        return '12'

def refatora_data(data):
    diamesano = data.split()
    mes = acha_mes(diamesano[2])

    data_formatada = diamesano[0]+'-'+mes+'-'+diamesano[4]

    return data_formatada

CC_costadoce = st.secrets['CC_COSTADOCE']
CC_fronteiraoeste = st.secrets['CC_FRONTEIRAOESTE']

email_alegrete = st.secrets['EMAIL_ALEGRETE']

email_bage = st.secrets['EMAIL_BAGE']

email_cachoeira = st.secrets['EMAIL_CACHOEIRA']

email_camaqua = st.secrets['EMAIL_CAMAQUA']

email_pedrito= st.secrets['EMAIL_PEDRITO']

email_pelotas = st.secrets['EMAIL_PELOTAS']

email_borja = st.secrets['EMAIL_BORJA']

email_gabriel= st.secrets['EMAIL_GABRIEL']

email_lourenco = st.secrets['EMAIL_LOURENCO']

email_gonzaga = st.secrets['EMAIL_GONZAGA']

email_uruguaiana= st.secrets['EMAIL_URUGUAIANA']


lista_alegrete = ['alegrete','manoel viana','quarai','sao francisco de assis']

lista_bage = ['acegua','bage','candiota','hulha negra']

lista_cachoeira = ['agudo','cacapava do sul','cachoeira do sul', 'candelaria','cerro branco', 'encruzilhada do sul','gramado xavier','herveiras',
                  'mato leitao','novo cabrais','pantano grande','paraiso do sul','passo do sobrado','rio pardo','santa cruz do sul',
                  'santana da boa vista','sinimbu','vale do sol','venancio aires','vera cruz']

lista_camaqua = ['camaqua','amaral ferrador','arambare','barra Do ribeiro','cerro grande do sul','chuvisca',
                  'cristal','dom feliciano','sentinela do sul','sertao santana','tapes']

lista_pedrito = ['dom pedrito','lavras do sul','santana do livramento']

lista_pelotas = ['pelotas','arroio grande','arroio do padre','cangucu','capao do leao','cerrito','chui','herval','jaguarao','morro redondo',
                  'pedras altas','pedro osorio','pinheiro machado','piratini','rio grande','santa vitoria do palmar','sao jose do norte','turucu']

lista_borja = ['garruchos','itacurubi','itaqui','macambara','santo antonio das missoes','sao borja']

lista_gonzaga = ['bossoroca','sao luiz gonzaga','caibate','campina das missoes','capao do cipo','cerro largo','dezesseis de novembro','guarani das missoes',
                  'mato queimado','pirapo','porto xavier','rolador','roque gonzales','salvador das missoes','santiago','sao nicolau','sao paulo das missoes',
                  'sao pedro do butia','unistalda','vitoria das missoes']

lista_uruguaiana = ['uruguaiana','barra do quarai']

lista_lourenco = ['sao lourenco do sul']

lista_gabriel = ['sao gabriel','rosario do sul','santa margarida do sul','vila nova do sul']

dict_cidades = {'São Borja':lista_borja,'Dom Pedrito':lista_pedrito,"São Luiz Gonzaga":lista_gonzaga,"Camaquã":lista_camaqua,'Alegrete':lista_alegrete,
                "Pelotas":lista_pelotas,"Cachoeira do Sul":lista_cachoeira,"Bage":lista_bage,'Uruguaiana':lista_uruguaiana, "São Lourenço do Sul":lista_lourenco,'São Gabriel':lista_gabriel}

def qual_loja(cidade):
    if cidade in lista_alegrete:
        return email_alegrete,CC_fronteiraoeste
    
    if cidade in lista_bage:
        return email_bage,CC_costadoce
    
    if cidade in lista_cachoeira:
        return email_cachoeira,CC_costadoce
    
    if cidade in lista_camaqua:
        return email_camaqua,CC_costadoce
    
    if cidade in lista_pedrito:
        return email_pedrito,CC_fronteiraoeste
    
    if cidade in lista_pelotas:
        return email_pelotas, CC_costadoce
    
    if cidade in lista_borja:
        return email_borja,CC_fronteiraoeste
    
    if cidade in lista_gabriel:
        return email_gabriel,CC_fronteiraoeste
    
    if cidade in lista_lourenco:
        return email_lourenco,CC_costadoce
    
    if cidade in lista_gonzaga:
        return email_gonzaga, CC_fronteiraoeste
    
    if cidade in lista_uruguaiana:
        return email_uruguaiana,CC_fronteiraoeste
    
    return 'erro'

def prepara_planilha_licencas(df, email, senha):
    df_lower = df

    df['Data de Início'] = df['Data de Início'].apply(refatora_data)

    df['Data de Início'] = pd.to_datetime(df['Data de Início'], format='%d-%m-%Y')

    df['Data Final'] = df['Data Final'].apply(refatora_data)
    agora = pd.Timestamp.now()

    df['Data Final'] = pd.to_datetime(df['Data Final'], format='%d-%m-%Y')

    df['Dias até o vencimento'] = (df['Data Final']-agora).dt.days+2
    df_lower['Cidade da Organização'] = df_lower['Cidade da Organização'].str.lower()
    df_lower['Cidade da Organização'] = df_lower['Cidade da Organização'].apply(remover_acentos)
    lista_cidades = df['Cidade da Organização'].unique().tolist()

    for chave, valor in dict_cidades.items():
        df_cidade = df_lower[df_lower['Cidade da Organização'].isin(valor)]
        if len(df_cidade) == 0:
            continue
        normalizada = remover_acentos(chave)
        normal_menor = normalizada.lower()
        dict_msg = {'Licença':[], 'Organização':[],'Equipamento':[],"Status":[],'Dias até o vencimento':[],'Início/Fim':[]}

        for index, row in df_cidade.iterrows():
            licenca = str(row['Nome da Licença'])+' - '+str(row['Número da Licença'])
            org = str(row['Nome do Cliente'])+' - '+str(row['OrgId'])
            equip = str(row['Modelo'])+' '+str(row['Nº de Série'])
            expirar = str(row['Dias até o vencimento'])+" dias"
            ini_fim = str(row['Data de Início'])+' '+str(row['Data Final'])

            if row['Dias até o vencimento']<0:
                dict_msg['Status'].append("Expirou")
            else:
                dict_msg['Status'].append("Ativo (Expira <7 dias)")

            dict_msg['Licença'].append(licenca)
            dict_msg['Organização'].append(org)
            dict_msg['Equipamento'].append(equip)
            dict_msg['Dias até o vencimento'].append(expirar)
            dict_msg['Início/Fim'].append(ini_fim)

        df_msg = pd.DataFrame(dict_msg)
        gerente_Email = Emailer(email,senha)
        destino,copia = qual_loja(normal_menor)

        assunto = "Testando email atuomático pra mais de uma pessoa - "
        corpo = f"""
        <html>
        <body>
            <p>Bom dia, espero que esteja bem.</p>
 
            <p>Escrevo para informar sobre a possibilidade de incremento de faturamento neste mês por meio da renovação de licenças. Abaixo compartilho imagem com as licenças e os respectivos equipamentos que as utilizam.</p>
            {df_msg.to_html(index=False)}
        </body>
        </html>
        """

        pop_up = gerente_Email.enviar_email_outlook(assunto=assunto,corpo_html=corpo,destinatario=destino,copia=copia,planilha=df_msg)

        if 'Email enviado com sucesso para' in pop_up:
            st.toast(pop_up)
        else:
            st.error(pop_up)
            break

st.title("Envio de email automatizado")

remetente = st.text_input("E-mail do remetente")
senha = st.text_input("Senha",type="password")
arquivo = st.file_uploader("Envie as licenças que estão por vencer", type=["xlsx"])
barra = st.progress(0)

if arquivo is not None:
    st.success("Arquivo carregado com sucesso!")
    df_uploaded = pd.read_excel(arquivo)
    if st.button("Envio das lincenças"):
        if '@alvorada-rs.com.br' in remetente:
            prepara_planilha_licencas(df_uploaded,remetente,senha)
        else:
            st.error("Email inválido")


# data = refatora_data('27 de fev. de 2024')

# print(data)

# df = pd.read_excel('c:\\Users\\GabrielHentschke-Alv\\OneDrive - Alvorada Sistemas Agrícolas Ltda\\Downloads\\Licenças_2026_02_13.xlsx')

#df_ativo = df[df["Status"]=='Ativo']

# prepara_planilha_licencas(df)

# prepara_planilha_licencas(df)
# teste = 'Sant\'Ana do Livramento'
# novo = remover_acentos(teste)
# print(novo)
# print(qual_loja(novo.lower()))