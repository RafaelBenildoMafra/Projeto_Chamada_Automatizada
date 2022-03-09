#código enviado apenas para os professores analizarem. Para rodar, é necessário ter todos os arquivos e caminhos iguais aos
#informados. Caso deseje os arquivos (de imagem, xml, etc), favor entrar em contato com os autores.

from tkinter import *
from tkinter import messagebox, ttk
import cv2
import datetime
import os
import numpy as np
import time
import pymysql
import win32com.client as win32
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

root = Tk()
root.iconbitmap('brasao_UFSC_vertical_sigla.ico')
root.geometry("490x560+610+153")
root.resizable(0, 0)
root.title('Sistema de Chamada Automatizada')

img_fundo = PhotoImage(file="imagens/fundo.png")
lab_fundo = Label(root, image=img_fundo)
lab_fundo.pack()

db = pymysql.connect(
    host='den1.mysql3.gear.host',
    user='rafaelmafradb',
    password='Ui00AkEsp~~2',
    database='rafaelmafradb',
)
mycursor = db.cursor()


def iniciar_aula():
    aba_aula = Tk()
    aba_aula.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    aba_aula.geometry("270x270")
    aba_aula.title("Iniciar Aula")
    aba_aula.configure(bg='#26abff')

    lista_aula = Tk()
    lista_aula.geometry("490x560+610+153")
    lista_aula.resizable(0, 0)
    lista_aula.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    lista_aula.title("Lista de Aulas")

    def aula_selecionada():
        escolha_aula = drop.get()
        messagebox.showwarning("AVISO", "AULA INICIADA PREPARANDO PARA CHAMADA!")
        aba_aula.destroy()
        lista_aula.destroy()
        # Camera 1 reconhecimento das pessoas
        camera1 = cv2.VideoCapture(0)
        # Camera 2 quantidade de pessoas
        camera2 = cv2.VideoCapture(1)
        nome = ""
        numerodepessoas = 0
        contagemloop = 0
        contagemif = 0
        mycursor.execute("DELETE from presenca")
        # Camera 1 reconhecimento das pessoas
        detectorFace = cv2.CascadeClassifier("F:\PI\haarcascade_frontalface_default.xml")
        reconhecedor = cv2.face.LBPHFaceRecognizer_create()
        reconhecedor.read("F:\PI\classificadorlbph.yml")
        # Camera 2 quantidade de pessoas
        classificadorVideoFace = cv2.CascadeClassifier('F:\PI\haarcascade_frontalface_default.xml')
        largura, altura = 220, 220
        font = cv2.FONT_HERSHEY_COMPLEX_SMALL

        while (True):
            conectado, imagem = camera1.read()  # Camera1
            imagemCinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
            facesDetectadas = detectorFace.detectMultiScale(imagemCinza,
                                                            scaleFactor=1.5,
                                                            minSize=(30, 30))
            conectado2, frame = camera2.read()  # camera2
            cinza = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            detecta = classificadorVideoFace.detectMultiScale(cinza, scaleFactor=1.1, minNeighbors=8, minSize=(25, 25))

            for (x, y, l, a) in facesDetectadas:

                # camera 1 parâmetros
                imagemFace = cv2.resize(imagemCinza[y:y + a, x:x + l], (largura, altura))
                cv2.rectangle(imagem, (x, y), (x + l, y + a), (0, 0, 255), 2)
                matricula, confianca = reconhecedor.predict(imagemFace)
                cv2.putText(imagem, nome, (x, y + (a + 30)), font, 2, (0, 0, 255))
                cv2.putText(imagem, str(confianca), (x, y + (a + 50)), font, 1, (0, 0, 255))

                if (confianca <= 35):
                    # print(matricula)
                    mycursor.execute("SELECT aluno FROM aluno WHERE matricula = %s", matricula)
                    pessoas = mycursor.fetchone()
                    nome = pessoas[0]
                    ja_foichamado = mycursor.execute(
                        "SELECT presente FROM Presenca WHERE matricula =" + "'" + str(matricula) + "'")
                    if (ja_foichamado != 1):
                        presente = 1
                        horarioentrada = datetime.datetime.now()
                        current_time = horarioentrada.strftime("%d/%m/%Y %H:%M:%S")
                        messagebox.showinfo("AVISO",
                                            "O aluno " + nome + " foi registrado na chamada com entrada às " + current_time)
                        mycursor.execute(
                            "INSERT INTO Presenca(matricula, aluno, presente, data) VALUES (%s, %s, %s, %s)",
                            (matricula, nome, presente, current_time))
                        db.commit()

            for (x, y, l, a) in detecta:
                # Câmera2parametros
                cv2.rectangle(frame, (x, y), (x + l, y + a), (255, 0, 0), 2)

                contador = str(detecta.shape[0])
                cv2.putText(frame, contador, (x + 10, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2, cv2.LINE_AA)

                cv2.putText(frame, "Quantidade de Faces: " + contador, (10, 450), cv2.FONT_HERSHEY_SIMPLEX, 1,
                            (0, 255, 0), 2, cv2.LINE_AA)
                contagemloop = contagemloop + 1
                if (contagemloop >= 200):
                    contagemif = contagemif + 1
                    numerodepessoas = numerodepessoas + detecta.shape[0]
                    contagemloop = 0

            cv2.imshow("Camera Reconhecimento", imagem)
            cv2.imshow("Camera contagem", frame)

            if cv2.waitKey(1) == ord('q'):
                messagebox.showinfo("AVISO", "CHAMADA ENCERRADA")
                if (contagemif == 0):
                    contagemif = 1
                mediadepessoas = numerodepessoas / contagemif
                horariochamada = datetime.datetime.now()
                datechamada = horariochamada.strftime("%d-%m-%Y")
                caminho = os.path.dirname('F:\PI\chamadas')
                pdf = canvas.Canvas(caminho + "chamada" + datechamada + ".pdf", pagesize=A4)
                mycursor.execute("SELECT * FROM presenca")
                lista_presenca = mycursor.fetchall()
                mycursor.execute("SELECT professor FROM aula WHERE materia =" + "'" + escolha_aula + "'")
                lista_professor = mycursor.fetchone()
                professor = lista_professor[0]
                pdf.setTitle('Lista de chamada ' + escolha_aula + ' ' + str(datetime.time()))
                # draw_my_ruler(pdf)
                pdf.setFont("Courier-Bold", 14)
                pdf.drawString(80, 660, 'Lista de chamada automatizada ' + escolha_aula)
                pdf.drawString(80, 640, 'Gerada automaticamente:  ' + datechamada)
                pdf.drawString(80, 620, 'Professor: ' + professor)
                pdf.setFont("Courier", 8)
                pdf.drawString(30, 50,
                               'Desenvolvido por: Rafael Benildo Mafra, Eduardo Davila e Leonardo Ambrosio para a matéria de Projeto Integrador')
                # pdf.line(160,670,420,670)
                pdf.drawInlineImage('ufsc_logo.jpeg', 100, 700)
                text = pdf.beginText(80, 580)
                text.setFont("Courier", 12)
                for presenca in lista_presenca:
                    text.textLine(str(presenca[0]) + '   ' + str(presenca[1]) + '   ' + str(presenca[3]))

                pdf.drawText(text)
                pdf.save()


                # criar a integração com o outlook
                outlook = win32.Dispatch('outlook.application')

                # criar um email
                email = outlook.CreateItem(0)

                # configurar as informações do seu e-mail
                email.To = "agostedu@gmail.com"
                email.Subject = "Lista de presença aula de Calculo do dia " + datechamada
                email.HTMLBody = f"""
                       <p>Olá Professor {professor},

                       <p>\nSegue em anexo o PDF da lista de chamada da aula de {escolha_aula} do dia {datechamada}</p>

                       <p>\nA média de alunos durante a aula foi de {mediadepessoas} alunos</p>

                       <p>\nE-mail automático enviado, favor não responder</p>
                       """

                anexo = caminho + "chamada" + datechamada + ".pdf"
                email.Attachments.Add(anexo)

                email.Send()
                messagebox.showinfo("AVISO", "PDF CRIADO E E-MAIL ENVIADO COM SUCESSO!")

                break
        camera1.release()
        camera2.release()
        cv2.destroyAllWindows()


    mycursor.execute("SELECT * FROM aula")
    aulas = mycursor.fetchall()
    scroll = Scrollbar(lista_aula)
    scroll.pack(side=RIGHT, fill=Y)
    listbox = Listbox(lista_aula, yscrollcommand=scroll.set)
    listbox.configure(bg='#FFFFF0', font="Verdana 8 bold", fg='black')
    for aula in aulas:
        listbox.insert(END, '')
        listbox.insert(END, ' Materia: ' + str(aula[0]))
        listbox.insert(END, ' Professor: ' + str(aula[1]))
        listbox.insert(END, '------------------------------------------------')
        listbox.pack(fill=BOTH, expand=1)
        scroll.config(command=listbox.yview())

    materia = []
    for aula in aulas:
        materia.append(aula[0])

    clicked = StringVar(aba_aula)
    clicked.set(materia[0])
    drop = ttk.Combobox(aba_aula, values = materia)
    drop.set(materia[0])
    drop.grid(row=2, column=1, pady=(2, 2))
    drop.place(width=150, height=30, x=20, y=50)

    valor_label = Label(aba_aula, text='SELECIONAR MATERIA', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    valor_label.grid(row=0, column=1, pady=(2, 2))

    bt_coletar = Button(aba_aula, bd=2, text="Iniciar Aula", command=aula_selecionada)
    bt_coletar.grid(row=4, column=0, pady=(2, 2))
    bt_coletar.place(width=200, height=30, x=20, y=100)



def exibir_alunos():
    exibir = Tk()
    exibir.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    exibir.geometry("490x560+610+153")
    exibir.resizable(0, 0)
    exibir.title('Alunos Cadastrados')

    mycursor.execute("SELECT * FROM aluno")
    lista_presenca = mycursor.fetchall()
    scroll = Scrollbar(exibir)
    scroll.pack(side=RIGHT, fill=Y)
    listbox = Listbox(exibir, yscrollcommand=scroll.set)
    listbox.configure(bg='#FFFFF0', font="Verdana 8 bold", fg='black')
    for presenca in lista_presenca:
        listbox.insert(END, '')
        listbox.insert(END, ' Matricula: ' + str(presenca[0]))
        listbox.insert(END, ' Aluno: ' + str(presenca[1]))
        listbox.insert(END, '------------------------------------------------')
        listbox.pack(fill=BOTH, expand=1)
        scroll.config(command=listbox.yview())


def cadastrar_aluno():
    aba_aluno = Tk()
    classificador = cv2.CascadeClassifier("F:\PI\haarcascade_frontalface_default.xml")
    camera = cv2.VideoCapture(0)
    aba_aluno.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    aba_aluno.geometry("270x270")
    aba_aluno.title("Cadastro Aluno")
    aba_aluno.configure(bg='#26abff')

    def matricula_aluno():
        matricula = en_matricula.get()
        nome = en_nome.get()

        ja_matriculado = mycursor.execute("SELECT matricula FROM aluno WHERE matricula =" + "'" + matricula + "'")

        if (ja_matriculado == 1):
            messagebox.showerror("ERRO", "ALUNO JA CADASTRADO NO SISTEMA")

        else:
            mycursor.execute("INSERT INTO Aluno(matricula, aluno) VALUES (%s, %s)", (matricula, nome))
            db.commit()
            messagebox.showinfo("AVISO", "ALUNO CADASTRADO COM SUCESSO")
            messagebox.showwarning("AVISO", "CAPTURANDO FOTOS. O PROCESSO LEVA CERCA DE 1 MINUTO!")
            aba_aluno.destroy()

            amostra = 1
            numero_amostras = 25  # validar número de fotos
            largura, altura = 220, 220

            while (True):
                conectado, imagem = camera.read()
                imagemcinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
                facesdetectadas = classificador.detectMultiScale(imagemcinza,
                                                                 scaleFactor=1.5,
                                                                 minSize=(150, 150))
                for (x, y, l, a) in facesdetectadas:
                    cv2.rectangle(imagem, (x, y), (x + l, y + a), (0, 0, 255), 2)
                    if cv2.waitKey(1):
                        imagemface = cv2.resize(imagemcinza[y:y + a, x: x + l], (largura, altura))
                        cv2.imwrite("F:\PI\Fotos\pessoa." + str(matricula) + "." + str(amostra) + ".jpg",
                                    imagemface)  ##conexão de servidor
                        #print("foto " + str(amostra) + " capturada")
                        if(amostra==12):
                            messagebox.showinfo("AVISO", "METADE DAS FOTOS FORAM CAPTURADAS. CONTINUE O PROCESSO")
                        amostra += 1
                        time.sleep(1)
                        # tirando as fotos

                cv2.imshow("Captura de fotos", imagem)
                cv2.waitKey(1)

                if (amostra >= numero_amostras):
                    break  # fazer retornar pra função inicial

            camera.release()
            cv2.destroyAllWindows()

            lbph = cv2.face.LBPHFaceRecognizer_create()

            def getImagemComId():
                caminhos = [os.path.join('F:\PI\Fotos', f) for f in os.listdir('F:\PI\Fotos')]
                faces = []
                ids = []
                for caminhoImagem in caminhos:
                    imagemface = cv2.cvtColor(cv2.imread(caminhoImagem), cv2.COLOR_BGR2GRAY)
                    matricula = int(os.path.split(caminhoImagem)[-1].split('.')[1])
                    ids.append(matricula)
                    faces.append(imagemface)

                return np.array(ids), faces

            ids, faces = getImagemComId()

            # print('treinando')

            lbph.train(faces, ids)
            lbph.write('F:\PI\classificadorlbph.yml')

            messagebox.showinfo("AVISO", "FOTOS TIRADAS E TREINAMENTO FINALIZADO!")



    nome_label = Label(aba_aluno, text='Matricula', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    nome_label.grid(row=1, column=0, pady=(2, 2))

    valor_label = Label(aba_aluno, text='Nome', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    valor_label.grid(row=3, column=0, pady=(2, 2))

    en_matricula = Entry(aba_aluno, bd=2, font=("Arial", 9))
    en_matricula.grid(row=2, column=0, pady=8, padx=15, ipadx=40, ipady=3)

    en_nome = Entry(aba_aluno, bd=2, font=("Arial", 9))
    en_nome.grid(row=4, column=0, pady=8, padx=15, ipadx=40, ipady=3)

    bt_coletar = Button(aba_aluno, bd=2, text="Cadastrar Aluno", command=matricula_aluno)
    bt_coletar.place(width=140, height=40, x=60, y=285)
    bt_coletar.grid(row=6, column=0, pady=30, padx=15, ipadx=40, ipady=3)


def remover_aluno():
    aba_aluno = Tk()
    aba_aluno.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    aba_aluno.geometry("270x270")
    aba_aluno.title("Remover Aluno")
    aba_aluno.configure(bg='#26abff')

    def removermatricula_aluno():
        matricula = en_matricula.get()

        try:
            mycursor.execute("SELECT * FROM aluno WHERE matricula =" + "'" + matricula + "'")

            resultado = mycursor.fetchall()

            messagebox.showinfo("AVISO", "ALUNO ENCONTRADO:\n" + str(resultado[0][1]) + '\n' + "matricula: " + str(
               resultado[0][0]))

            aba_aluno.destroy()

            messagebox.showinfo("AVISO", "ALUNO REMOVIDO COM SUCESSO")
        except:
            messagebox.showerror("AVISO", "ALUNO NÃO ENCONTRADO, TENTE NOVAMENTE!")

        mycursor.execute("DELETE FROM aluno WHERE matricula =" + "'" + matricula + "'")
        db.commit()



    nome_label = Label(aba_aluno, text='Matricula', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    nome_label.grid(row=1, column=0, pady=(2, 2))

    en_matricula = Entry(aba_aluno, bd=2, font=("Arial", 9))
    en_matricula.grid(row=2, column=0, pady=8, padx=15, ipadx=40, ipady=3)

    bt_coletar = Button(aba_aluno, bd=2, text="Remover Aluno", command=removermatricula_aluno)
    bt_coletar.place(width=140, height=40, x=60, y=285)
    bt_coletar.grid(row=6, column=0, pady=30, padx=15, ipadx=40, ipady=3)


def cadastrar_aula():
    aba_aluno = Tk()
    aba_aluno.iconbitmap('brasao_UFSC_vertical_sigla.ico')
    aba_aluno.geometry("270x270")
    aba_aluno.title("Cadastrar Aula")
    aba_aluno.configure(bg='#26abff')

    def cadastro():
        materia = en_materia.get()
        professor = en_professor.get()

        mycursor.execute("INSERT INTO Aula(materia, professor) VALUES (%s, %s)", (materia, professor))
        db.commit()

        aba_aluno.destroy()

        messagebox.showinfo("AVISO", "MATÉRIA CADASTRADA COM SUCESSO!")

    nome_label = Label(aba_aluno, text='Nome da matéria', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    nome_label.grid(row=1, column=0, pady=(2, 2))

    valor_label = Label(aba_aluno, text='Professor', bg='#26abff', font="Arial 14 bold", fg='#00008B')
    valor_label.grid(row=3, column=0, pady=(2, 2))

    en_materia = Entry(aba_aluno, bd=2, font=("Arial", 9))
    en_materia.grid(row=2, column=0, pady=8, padx=15, ipadx=40, ipady=3)

    en_professor = Entry(aba_aluno, bd=2, font=("Arial", 9))
    en_professor.grid(row=4, column=0, pady=8, padx=15, ipadx=40, ipady=3)

    bt_coletar = Button(aba_aluno, bd=2, text="Cadastrar Aula", command=cadastro)
    bt_coletar.place(width=140, height=40, x=60, y=285)
    bt_coletar.grid(row=6, column=0, pady=30, padx=15, ipadx=40, ipady=3)



# Criação de botões
bt_aluno = Button(root, bd=2, command=iniciar_aula, text="Iniciar Aula", font="Arial 9 ", fg='black')
bt_aluno.place(width=140, height=40, x=180, y=165)

bt_aluno1 = Button(root, bd=2, command=cadastrar_aluno, text="Cadastrar Aluno", font="Arial 9 ", fg='black')
bt_aluno1.place(width=140, height=40, x=100, y=225)

bt_aluno2 = Button(root, bd=2, command=exibir_alunos, text="Exibir Alunos", font="Arial 9 ", fg='black')
bt_aluno2.place(width=140, height=40, x=260, y=225)

bt_professor = Button(root, bd=2, command=cadastrar_aula, text="Cadastrar Aula", font="Arial 9 ", fg='black')
bt_professor.place(width=140, height=40, x=100, y=285)

bt_professor1 = Button(root, bd=2, command=remover_aluno, text="Remover Aluno", font="Arial 9 ",
                       fg='black')
bt_professor1.place(width=140, height=40, x=260, y=285)

root.mainloop()
