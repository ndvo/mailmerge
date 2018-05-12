#! /usr/bin/python
# -*- coding: utf-8 -*-
import codecs
import smtplib
import re
import os
import sys
import time
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import Encoders
from datetime import datetime
import readline
from unidecode import unidecode

readline.parse_and_bind("tab: complete")

j = os.path.join
prompt=">>> "
#para cada variavel da mala direta, registre um indice na coluna no seguinte formato:
# 'variavel':numero da coluna
colunas = {
        'mail':0,
	}


checar = True
raw_input("""
~*~*~ Enviar emails em massa ~*~*~

Este script irá enviar uma mala-direta a partir de arquivos salvos em um determinado diretório.

Instruções:
- O script fará as perguntas necessárias para o envio dos emails
- As respostas padrão aparecem entre colchetes
- Para cancelar o script, a qualquer momento pressione Ctrl+c
-Use a tecla Tab para que o script tente autocompletar as respostas

---Pressione qualquer tecla para continuar---""")

remetente = raw_input("Informe o remetente [nao-responda@planejamento.gov.br]\n"+prompt) or 'nao-responda@planejamento.gov.br'
print "Remetente = ", remetente
incluirimagens = raw_input("""
Você pode incluir imagens no seu email.

As imagens devem estar no mesmo diretório da mensagem
e devem ser referenciadas no HTML com o padrão cid:NOME-DA-IMAGEM-SEM-EXTENSAO.
A imagem deve estar no formato .png
Você gostaria de incluir as imagens disponíveis no diretório como anexo? (S/N) [S]\n"""+prompt) or "S"

if incluirimagens in ("S","s","Sim","sim","Y","y"):
	incluirimagens=True
	print "Incluir imagens: SIM"
else:
	print "Incluir imagens: NÃO"

raw_input("""

Preparando o diretório.

Para enviar emails em massa crie uma pasta e salve nela os seguintes arquivos:
- assunto.txt (opcional: arquivo com o texto do assunto)
- base.csv (arquivo com os destinatários)
- teste.csv (arquivo com destinatários de teste)
- excluir.csv (arquivo com destinatários a excluir do envio)
- mensagem.txt (corpo do email, em formato TXT)
- mensagem.html (corpo do email, em formato HTML)
- anexo.pdf (anexo a ser enviado com o email)

---Pressione qualquer tecla para continuar---""")

naoenviar = "excluir.csv"


while checar:
	checar = False
	diretorio = raw_input("Informe o diretorio onde estão os arquivos [teste]\n"+prompt) or "teste"
	print "Diretório = ",diretorio
	print "Verificando os arquivos..."
	csvfile = j(diretorio, 'base.csv')
	testefile = j(diretorio, 'teste.csv')
	arquivo_txt = j(diretorio, 'mensagem.txt')
	arquivo_html= j(diretorio, 'mensagem.html')
	anexo=j(diretorio, "anexo.pdf")
	excluir= j(diretorio, naoenviar)
	if not os.path.isfile(anexo):
		anexo = False
		print "O arquivo anexo não existe. Não será encaminhado nenhum anexo."
	for i in (csvfile, arquivo_txt, arquivo_html, testefile, excluir):
		if not os.path.isfile(i):
			print "O arquivo "+i+" não existe."
			checar = True
			if i == excluir:
				open(i,'a').close()
				print "Foi criado um arquivo "+i+" vazio."
 	if checar:
		print raw_input("Verifique os arquivos para que possamos continuar.")

assunto="Mensagem de testes"
if ('assunto.txt' in os.listdir(diretorio)):
	assunto = open(j(diretorio,'assunto.txt')).read()
assunto = raw_input("Informe o assunto do email ["+assunto+"]:\n"+prompt) or assunto
print "Assunto = ", assunto
		
def enviar_email(remetente, assunto, destinatarios):
        # Cria um container para a mensagem.
	msg = MIMEMultipart('related')
        alt = MIMEMultipart('alternative')
        #msg['Subject'] = assunto #MIMEText(assunto,'plain','utf-8')
	msg['Subject']=Header(assunto, 'utf-8')
	#msg['Content-Type'] = "text/html; charset=utf-8"
        msg['From'] = remetente
	msg.preamble =  'This is a multi-part message in MIME format.'
	msg.attach(alt)
        # Cria o corpo da mensagem (uma versão em HTML e outra em TXT).
        text = open(arquivo_txt).read()
        html = open(arquivo_html).read()
        # Registra os tipos de ambas as partes.
        part1 = MIMEText(text, 'plain',"utf-8")
        part2 = MIMEText(html, 'html',"utf-8")
        # Anexa as partes ao container
        # O segundo a ser anexado é o preferencial
        alt.attach(part1)
        alt.attach(part2)
	if anexo:
		mimeanexo = MIMEBase('application','pdf')
		mimeanexo.set_payload(file(anexo).read())
		Encoders.encode_base64(mimeanexo)
		mimeanexo.add_header('Content-Disposition','attachment',filename=anexo)
		msg.attach(mimeanexo)
	#Incluir as imagens como anexos
	if incluirimagens:
		listaimagens = [i for i in os.listdir(j(diretorio)) if i[-4:]==".png"]
		for i in listaimagens:
			# This example assumes the image is in the current directory
			fp = open(j(diretorio,i), 'rb')
			msgImage = MIMEImage(fp.read())
			fp.close()
			# Define the image's ID as referenced above
			msgImage.add_header('Content-ID', '<'+i[:-4]+'>')
			msg.attach(msgImage)
	try:
		# Envia a mensagem.
		s = smtplib.SMTP('localhost')
		#if destinatario not in enviados:
		s.sendmail(remetente, destinatarios, msg.as_string())
		s.quit()
		logar = now.strftime("%d-%m-%Y %H:%M")+" OK    - "+assunto+" "+str(destinatarios)+"\n"
	except Exception as error:
		logar = now.strftime("%d-%m-%Y %H:%M")+" Error - "+assunto+" "+str(error)+"\n"
	with open(j(diretorio, 'log'), 'a') as enviadoslog:
		enviadoslog.write(logar)

def corretor(linha):
	correcoes = {
		'\.\.':'.',
		'@\.':'@',
		'[\.@]@':'@',
		'\.[bB]$':'.br',
		'[^a-zA-Z]+$':'',
		'receita.fazenda$':'receita.fazenda.gov.br',
		'RECEITA.FAZENDA$':'RECEITA.FAZENDA.GOV.BR',
		'\s*':'',
		'\s*':'',
		'\s*':'',
	}
	#aplica as correcoes
	for padrao,substituto in correcoes.iteritems():
		linha = re.sub(padrao,substituto, linha)
	if isinstance(linha, unicode):
		linha = unidecode(linha)
	#remoção de espaços em branco
	linha=linha.strip()
	return linha

def get_lista_exclusao():
	global naoenviar	
	lista_excluir = {}
	with (open(j(diretorio, naoenviar),'rb')) as exclusao:
		for linha in sorted(exclusao.readlines()):
			if linha[0] not in lista_excluir:
				lista_excluir[linha[0]] = []
			lista_excluir[linha[0]].append(linha)
	return lista_excluir

def sanea_base(base):
	excluir = get_lista_exclusao()
	print "\n"+str(sum([len(v) for k,v in excluir.iteritems()]))+" emails constam da lista de exclusão.\n"
	with (open(base,'rb')) as base:
		#Regras para email
		#Nome de usuário pode ter letras, números, sinal de mais, underscore, hífen, percentual, ponto
		#Nome de usuário pode começar com letras, números, undescore
		#Nome de domínio pode ter letras, números, hífen
		#Nome de domínio pode começar com letra, número
		mailre = re.compile('^[a-zA-Z0-9_][a-zA-Z0-9._%+-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*[a-zA-Z0-9]\.[a-zA-Z]{2,6}$',re.UNICODE )
		previous = ""
		aenviar = {}
		erros = 0
		for linha in sorted(base.readlines()):
			if len(linha)>0:
				letra=linha[0]
				linha=corretor(linha)
				if linha == previous:
					print "Erro: Email repetido: ", linha
				else:
					endereco = linha
					letra=endereco[0]
					if letra not in excluir.keys() or endereco not in excluir[letra]:
						if mailre.match(endereco):
							dominio = endereco.split('@')[-1]
							if dominio not in aenviar.keys():
								aenviar[dominio]=[]
							aenviar[dominio].append(endereco)
						#print "Enviando para "+linha[0]
					else:
						print "Erro: Email mal formatado: ", linha
			previous = linha
	return aenviar

enviados = []
pacote = int(raw_input("Os emails serão enviados em pacotes. Quantos emails deve conter cada pacote? [50]\n"+prompt) or 50)
start_time = datetime.now()
simulacao = True

while simulacao:
	atual = 0
	simulacao = False if raw_input("O script vai simular o envio, mandando as mensagens para a base de testes. Caso você queira efetivamente enviar, digite Enviar\n"+prompt) in ("Enviar","enviar") else True
	if simulacao:
		print raw_input("Você optou pela simulação. Será feito o saneamento da base de testes e em seguida os emails serão enviados para a base de testes. Pressione qualquer tecla para continuar.\n")
		#ao sanear a base ele 
		aenviar = sanea_base(testefile)
		contador = str(sum([len(v) for k,v in aenviar.iteritems()]))
	else:
		print raw_input("ATENÇÃO: Você optou pelo envio. Pressione qualquer tecla para continuar.\n") 
		print raw_input("A lista de emails agora será saneada e filtrada.\nEsta etapa pode demorar.\nPressione enter para continuar e aguarde até que o saneamento da base seja concluído.")
		aenviar = sanea_base(csvfile)
		contador = sum([len(v) for k,v in aenviar.iteritems()])
	print ("Serão enviados "+str(contador)+" emails divididos em pacotes de "+str(pacote)+" emails.\n")
	print "Início:"
	now = datetime.now()
	print now.strftime("%d-%m-%Y %H:%M")
	for dominio,enderecos in aenviar.iteritems():
		#print enderecos
		for i in [enderecos[a:a+pacote] for a in xrange(0, len(enderecos), pacote)]:
			if simulacao:
				enviar_email(remetente, '[TESTE] '+assunto,  i )
			else:
				enviar_email(remetente, assunto,  i )
		numerodominio=len(enderecos)
		atual+=numerodominio
		elapsed = str(datetime.now() - start_time )
		percentual = "{:.2f}".format((float(atual)/float(contador))*100)
		sys.stdout.write("\r "+percentual+"% --- "+elapsed+" --- "+dominio+" com "+str(numerodominio)+" emails.               ")
		#sys.stdout.flush()
		#print "\n"
	if simulacao:
		print "\nSimulação concluída.\n"
	else:
		print "\nForam enviados "+str(contador)+" emails.\n"
		print "Término:"
		now = datetime.now()
		print now.strftime("%d-%m-%Y %H:%M")
		raw_input("Pressione qualquer tecla para encerrar.\n")
		exit()
