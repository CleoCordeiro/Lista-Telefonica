from os import system
import re
import pandas as pd
import msvcrt
import time

'''
Aluno: Cléo Maia Cordeiro
Curso: SISTEMAS DE INFORMACAO/CCAST - Castanhal - N
'''

arq = 'Lista_de_Contatos.xlsx'


def linha(tam=80):
    return '-' * tam


def titulo(txt):
    print(linha())
    print(f'\033[32m{txt.center(80)}\033[m')
    print(linha())


def menu(lista):
    c = 1
    for item in lista:
        print(f"\033[96m[{c}] - {item}\033[m")
        c += 1
    print(linha())
    opcao = leiaint('Digite a Opção Desejada: ')
    return opcao


def menulateral(lista):
    c = 1
    menu = ''
    for item in lista:
        menu += f"[{c}] - {item}   "
        c += 1
    print('')
    print(f'\033[96m{menu.center(80)}\033[m')
    print(linha())
    opcao = leiaint('Digite a Opção Desejada: ')
    return opcao


def leiaint(msg):
    while True:
        try:
            n = int(input(msg))
        except (ValueError, TypeError):
            print('\033[31mERRO: Por Favor, Digite um Número Inteiro Válido.\033[m')
        else:
            return n


def telefonevalidador(msg):
    while True:
        try:
            n = str(input(msg))
        except (ValueError, TypeError):
            print('\033[31mERRO: Por Favor, Digite um Número de Telefone Válido.\033[m')
        else:
            return n


def verificaseexistedb(arquivo):
    try:
        a = pd.read_excel(arq)
    except FileNotFoundError:
        return False
    else:
        return True


def criardb(arquivo):
    try:
        a = pd.DataFrame(columns=['Nome', 'Email', 'Telefone'])
        a.to_excel(arquivo, index=False)
    except:
        print('Não Foi Possível Criar o Arquivo de Database')


def atualizardb():
    global Lista_de_Contatos_df, arq
    Lista_de_Contatos_df.to_excel(arq, index=False)


def cadastrar():
    global Lista_de_Contatos_df, arq
    system('cls')
    titulo('CADASTRAR NOVO CONTATO')
    print('Digite o Nome do Contato')
    while True:
        nome = str(input('Nome: '))
        nomecompleto = nome.split(' ')
        nomecheckexist = Lista_de_Contatos_df.query('Nome.str.lower() == @nome.lower()', engine='python').head()
        if nome == '':
            print('Não é Possível Criar um Nome Vazio')
        elif not re.match(r"^[^0-9_!¡?÷?¿/\\+=@#$%ˆ&*(){}|~<>;:[\]]{2,}$", nome):
            print('Digite um Nome Completo Válido')

        elif len(nomecompleto) < 2:
            print('Digite um Nome Completo Válido')

        elif not nomecheckexist.empty:
            print("Esse Nome já Está Cadastrado")
        else:
            break
    print('Digite o Email do Contato')
    while True:
        email = str(input('Email: '))
        emailcheckexist = Lista_de_Contatos_df.query('Email.str.lower() == @email.lower()', engine='python').head()
        if email == '':
            print('Não é Possível Criar um Nome Vazio')
        elif not re.match(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email):
            print("Por Favor Digite um Email Válido")
        elif not emailcheckexist.empty:
            print("Esse email já Está Cadastrado")
        else:
            break
    print('Digite o Telefone do Contato')
    while True:
        telefone = input('Telefone: ')
        telefonecheckexist = Lista_de_Contatos_df.query('Telefone.str.lower() == @telefone.lower()',
                                                        engine='python').head()

        """Regex para verificação se o número de telefone é válido
        if not re.match(r'^(\(?\d{2}\)?)?\s?\d{4,5}-?\d{4,5}', telefone):
        Reduzi o numero de opções para aceitar apenas com ddd + numero
        exemplos de entradas válidas:
        91 999999999 or  91999999999 or 999999999"""

        if not re.match(r'^(\d{2})?\s?\d{4,5}\d{4,5}', telefone):
            print('Por Favor, Digite um Número de Telefone Válido')
        elif not telefonecheckexist.empty:
            print("Esse Telefone já Está Cadastrado")
        else:
            break

    cadastrar = pd.DataFrame({"Nome": [nome],
                              "Email": [email],
                              "Telefone": [telefone]
                              }, )
    Lista_de_Contatos_df = Lista_de_Contatos_df.append(cadastrar, ignore_index=True)
    try:
        atualizardb()
        system('cls')
        titulo('Contato Cadastrado com Sucesso')
        print(Lista_de_Contatos_df.iloc[-1:])
        system('pause')
    except:
        print('Não foi Possível Cadastrado o Novo Contato')
        Lista_de_Contatos_df = Lista_de_Contatos_df.remove(cadastrar, ignore_index=True)


def pesquisar():
    global Lista_de_Contatos_df
    pesquisa = ''
    while True:
        print(f'Digite o Nome, Email ou Telefone que Deseja Pesquisar: {pesquisa}')

        keypressed = msvcrt.getch().decode("utf-8", 'ignore')

        system('cls')
        titulo('PESQUISAR CONTATO')
        if keypressed == '\r':
            break
        elif keypressed == '\x08':
            pesquisa = pesquisa[: -1]
        elif keypressed == '\x03':
            exit()
        else:
            pesquisa = pesquisa + keypressed
            time.sleep(0.1)
            print(pesquisa)
            #pesquisa = str(input('Digite o Nome, Email ou Telefone que Deseja Pesquisar: '))
            #Usei a pesquisa query do pandas com contains que é quase que a mesma cosa do like em sql
            #Para passar uma variavel usa-se o @variavel

        pesquisa.strip('space')
        print(Lista_de_Contatos_df.query('Nome.str.contains(@pesquisa, case=False) |'
                                           ' Email.str.contains(@pesquisa, case=False) |'
                                           'Telefone.str.contains(@pesquisa, case=False)'
                                           , engine='python').head())
    while True:
        opcao = menulateral(
            ['Fazer Nova Busca', 'Menu Principal', 'Sair'])
        if opcao == 1:
            pesquisar()
            break
        elif opcao == 2:
            break
        elif opcao == 3:
            exit()
        else:
            print('Digite uma Opção Válida')

def editar():
    while True:
        global Lista_de_Contatos_df
        system('cls')
        titulo('EDITAR CONTATOS')
        print('Será Possível Editar um Contato por Meio do Email do Contato a ser Editado')
        email = str(input('Email do Contato que Deseja Editar: '))
        # Usei a pesquisa query do pandas com contains que é quase que a mesma cosa do like em sql
        # Para passar uma variável usa-se o @variável

        indice = (Lista_de_Contatos_df.index[Lista_de_Contatos_df['Email'] == email].tolist())
        if indice:
            print('Contato a ser Editado')
            print(Lista_de_Contatos_df.loc[indice])
            editando = False
            while not editando:
                opcao = menulateral(
                    ['Confirmar', 'Cancelar'])
                if opcao == 1:
                    editando = True
                    print('Você Pode Editar Todos os Dados ou Selecionar o Dado a ser Editado')
                    while True:
                        opcao = menulateral(
                            ['Editar Todos os Dados', 'Editar Nome', 'Editar Email', 'Editar Telefone', 'Cancelar'])
                        if opcao == 1:
                            print('Digite o Novo Nome do Contato')
                            while True:
                                nome = str(input('Nome: '))
                                nomecompleto = nome.split(' ')
                                nomecheckexist = Lista_de_Contatos_df.query('Nome.str.lower() == @nome.lower()',
                                                                            engine='python').head()

                                if nome == '':
                                    print('Não é Possível Criar um Nome Vazio')
                                elif not re.match(r"^[^0-9_!¡?÷?¿/\\+=@#$%ˆ&*(){}|~<>;:[\]]{2,}$", nome):
                                    print('Digite um Nome Completo Válido')

                                elif len(nomecompleto) < 2:
                                    print('Digite um Nome Completo Válido')
                                elif nome == Lista_de_Contatos_df.loc[indice[0], ['Nome'][0]]:
                                    break
                                elif not nomecheckexist.empty:
                                    print("Esse Nome já Está Cadastrado")
                                else:
                                    break
                            print('Digite o Novo Email do Contato')
                            while True:
                                email = str(input('Email: '))
                                emailcheckexist = Lista_de_Contatos_df.query('Email.str.lower() == @email.lower()',
                                                                             engine='python').head()
                                if email == '':
                                    print('Não é Possível Criar um Nome Vazio')
                                elif not re.match(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email):
                                    print("Por Favor Digite um Email Válido")
                                elif email == Lista_de_Contatos_df.loc[indice[0], ['Email'][0]]:
                                    break
                                elif not emailcheckexist.empty:
                                    print("Esse email já Está Cadastrado")
                                else:
                                    break
                            print('Digite o Novo Telefone do Contato')
                            while True:
                                telefone = input('Telefone: ')
                                telefonecheckexist = Lista_de_Contatos_df.query(
                                    'Telefone.str.lower() == @telefone.lower()', engine='python').head()
                                if not re.match(r'^(\d{2})?\s?\d{4,5}\d{4,5}', telefone):
                                    print('Por Favor, Digite um Número de Telefone Válido')
                                elif telefone == Lista_de_Contatos_df.loc[indice[0], ['Telefone'][0]]:
                                    break
                                elif not telefonecheckexist.empty:
                                    print("Esse Telefone já Está Cadastrado")
                                else:
                                    break

                            Lista_de_Contatos_df.loc[indice, 'Nome'] = nome
                            Lista_de_Contatos_df.loc[indice, 'Email'] = email
                            Lista_de_Contatos_df.loc[indice, 'Telefone'] = telefone
                            atualizardb()
                            print(f'\033[32mContato Editado com Sucesso\033[m')
                            print(Lista_de_Contatos_df.loc[indice])
                            system('pause')
                            break

                        elif opcao == 2:
                            while True:
                                print('Digite o Novo Nome do Contato')
                                nome = str(input('Nome: '))
                                nomecompleto = nome.split(' ')
                                nomecheckexist = Lista_de_Contatos_df.query('Nome.str.lower() == @nome.lower()',
                                                                            engine='python').head()
                                if nome == '':
                                    print('Não é Possível Criar um Nome Vazio')
                                elif not re.match(r"^[^0-9_!¡?÷?¿/\\+=@#$%ˆ&*(){}|~<>;:[\]]{2,}$", nome):
                                    print('Digite um Nome Completo Válido')

                                elif len(nomecompleto) < 2:
                                    print('Digite um Nome Completo Válido')
                                elif nome == Lista_de_Contatos_df.loc[indice[0], ['Nome'][0]]:
                                    print('Você Digitou o Mesmo Nome Cadastrado No Contato')
                                elif not nomecheckexist.empty:
                                    print("Esse Nome já Está Cadastrado")
                                else:
                                    break
                            Lista_de_Contatos_df.loc[indice, 'Nome'] = nome
                            atualizardb()
                            print(f'\033[32mContato Editado com Sucesso\033[m')
                            print(Lista_de_Contatos_df.loc[indice])
                            system('pause')
                            break

                        elif opcao == 3:
                            while True:
                                print('Digite o Novo Email do Contato')
                                email = str(input('Email: '))
                                emailcheckexist = Lista_de_Contatos_df.query('Email.str.lower() == @email.lower()',
                                                                             engine='python').head()
                                if email == '':
                                    print('Não é Possível Criar um Nome Vazio')
                                elif not re.match(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email):
                                    print("Por Favor Digite um Email Válido")
                                elif email == Lista_de_Contatos_df.loc[indice[0], ['Email'][0]]:
                                    print('Você Digitou o Mesmo Email Cadastrado No Contato')
                                elif not emailcheckexist.empty:
                                    print("Esse email já Está Cadastrado")
                                else:
                                    break

                            Lista_de_Contatos_df.loc[indice, 'Email'] = email
                            atualizardb()
                            print(f'\033[32mContato Editado com Sucesso\033[m')
                            print(Lista_de_Contatos_df.loc[indice])
                            system('pause')
                            break

                        elif opcao == 4:
                            while True:
                                print('Digite o Novo Telefone do Contato')
                                telefone = input('Telefone: ')
                                telefonecheckexist = Lista_de_Contatos_df.query(
                                    'Telefone.str.lower() == @telefone.lower()', engine='python').head()
                                if not re.match(r'^(\d{2})?\s?\d{4,5}\d{4,5}', telefone):
                                    print('Por Favor, Digite um Número de Telefone Válido')
                                elif telefone == Lista_de_Contatos_df.loc[indice[0], ['Telefone'][0]]:
                                    print('Você Digitou o Mesmo Telefone Cadastrado No Contato')
                                elif not telefonecheckexist.empty:
                                    print("Esse Telefone já Está Cadastrado")
                                else:
                                    break
                            Lista_de_Contatos_df.loc[indice, 'Telefone'] = telefone
                            atualizardb()
                            print(f'\033[32mContato Editado com Sucesso\033[m')
                            print(Lista_de_Contatos_df.loc[indice])
                            system('pause')
                            break

                elif opcao == 2:
                    break
                else:
                    print('Escolha uma das Alternativas Válida')

            opcao = menulateral(
                ['Editar Outro Contato', 'Menu Principal', 'Sair'])
            if opcao == 1:
                continue
            elif opcao == 2:
                break
            elif opcao == 3:
                exit()
            else:
                print('Escolha uma das Alternativas Válida')
        else:
            print('Nenhum Contato Encontrado Tente Novamente')


def excluir():
    global Lista_de_Contatos_df
    system('cls')
    titulo('EXCLUIR CONTATO')
    print('Será Possível Excluir um Contato por Meio do Email do Contato a ser Excluído')
    while True:
        email = str(input('Email do Contato que Deseja Excluir: '))
        # Usei a pesquisa query do pandas com contains que é quase que a mesma cosa do like em sql
        # Para passar uma variável usa-se o @variável

        indice = (Lista_de_Contatos_df.index[Lista_de_Contatos_df['Email'] == email].tolist())
        if indice:
            print('Contato a ser Excluído')
            print(Lista_de_Contatos_df.loc[indice])
            while True:
                opcao = menulateral(
                    ['Confirmar', 'Cancelar'])
                if opcao == 1:
                    Lista_de_Contatos_df = Lista_de_Contatos_df.drop(indice)
                    atualizardb()
                    print(f'\033[32mContato Excluído com Sucesso\033[m')
                    system('pause')
                    break
                elif opcao == 2:
                    break
                else:
                    print('Escolha uma das Alternativas Válida')

            opcao = menulateral(
                ['Excluir Outro Contato', 'Menu Principal', 'Sair'])
            if opcao == 1:
                continue
            elif opcao == 2:
                break
            elif opcao == 3:
                exit()
            else:
                print('Escolha uma das Alternativas Válida')
        else:
            print('Nenhum Contato Encontrado Tente Novamente')


def listarcontatos():
    global Lista_de_Contatos_df
    system('cls')
    page = 0
    maxpageelements = 5
    numerodecontatos = int(len(Lista_de_Contatos_df.index) - 2)
    lastpage = int(numerodecontatos / maxpageelements)
    while True:
        system('cls')
        titulo('LISTA DE CONTATOS')

        print(Lista_de_Contatos_df.iloc[(page * maxpageelements):(page * maxpageelements) + maxpageelements])
        if page > 0:
            print(f'\033[33m\n\tPágina Atual[{page}]\t\tPágina Anterior[{page - 1}]\t\tÚltima Página[{lastpage}]\033[m')
        else:
            print(f'\033[33m\n\tPágina Atual[{page + 1}]\t\tÚltima Página[{lastpage}]\033[m')
        # print(Lista_de_Contatos_df.head(page * maxpageelements,maxpageelements))
        opcao = menulateral(
            ['Avançar', 'Voltar', 'Menu Principal', 'Sair'])
        if opcao == 1:
            if lastpage > page:
                print(f'Página atual {page}')
                print(f'Última Página {lastpage}')
                page += 1

            else:
                print('Você Chegou Na Última Página')
                system('pause')
        elif opcao == 2:
            if page >= 1:
                page -= 1
            else:
                print('Você Chegou na Primeira Página')
                system('pause')
        elif opcao == 3:
            break
        elif opcao == 4:
            exit()
        else:
            print('Digite uma Opção Válida')
            system('pause')


if not verificaseexistedb(arq):
    criardb(arq)

Lista_de_Contatos_df = pd.read_excel(arq)

while True:
    system('cls')
    titulo('LISTA TELEFÔNICA')
    opcao = menu(
        ['Cadastrar Novo Contato', 'Buscar Contato', 'Editar Contato', 'Excluir Contato', 'Listar Contatos', 'Sair'])

    if opcao == 1:
        cadastrar()
    elif opcao == 2:
        pesquisar()
    elif opcao == 3:
        editar()
    elif opcao == 4:
        excluir()
    elif opcao == 5:
        listarcontatos()
    elif opcao == 6:
        exit()
    else:
        print('A Opção Digitada é Inválida')
