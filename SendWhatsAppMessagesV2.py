# %% [markdown]
# # Send WhatsApp Web Messages From Excel With Images

# %% [markdown]
# This program send messages via WhatsApp Web with images  
# The messages must be stored in an Excel file, and mus contain the following columns:  
# 
# CLIENTE: Name of destinatary  
# TELEFONE: Phone of destinatary  
# MENSAGEM: Message to be sent  
# 
# Other columns can be present such as name, address, etc, so, by using Excel text concatenation formulae, to send highly personalized messages, including special characters, icons, emoticons, links, etc.  
# 
# With messages, the program send the images selected (jpg, png, or gif).  
# 
# Notes:  
#  - The program works only in Google Chrome  
#  - The program waits a random time betweeen messages to avoid WhatsApp to detect automation.  
#  - The program displays a scrolling text, showing the historical of messages with sucess or fail.  
#  - The program try to send the message and the images, if there is an error, jumps to the next one.  
#  - At the end, the program saves an Excel file with same fields (Cliente, Telefone and Mensagem) and adds a column with sucess (and date of message) or fail.  
#  - The program is a quite slow, to allow Google Chrome to perform operations.
# 
# 

# %% [markdown]
# Libraries

# %%
# import libraries

# Basic Tkinter
import tkinter as tk
from tkinter import filedialog as fd
import tkinter.scrolledtext as st

# PIL to show images on Tkinter
from PIL import Image, ImageTk

# Pandas
import pandas as pd

# Just to get image file name from full path
from pathlib import Path

# Time to allow program wait few seconds during Chrome operations
import time

# To allow randomic waiting times (important to avoid Whatsapp account blocking)
import random

# Datetime to store current date of messages sent
# from datetime import date
import datetime as dt

# necessary libraries for Chrome operations:
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

# modified 29-oct-23
from selenium.webdriver.chrome.service import Service

# pip install webdriver_manager
# This librari updates automatically the Browser Manger (in this case, Chrome)
from webdriver_manager.chrome import ChromeDriverManager

# Necessary to convert messages from ASCII text into URL aceptable addresses (convert special characters, spaces, etc)
import urllib

# %% [markdown]
# Global Variables

# %%
# list of images to send
imgs_path = []

# %% [markdown]
# Functions

# %% [markdown]
# https://stackoverflow.com/questions/74214619/how-to-use-tkinter-after-method-to-delay-a-loop-instead-time-sleep/74215342?noredirect=1#comment131053675_74215342

# %%
def tksleep(t):
    # Function to delay process for t seconds
    # emulating time.sleep(seconds)
    # Thanks to link above
    ms = int(t*1000)
    root = tk._get_default_root()
    var = tk.IntVar(root)
    root.after(ms, lambda: var.set(1))
    root.wait_variable(var)

# %%
def show_imgs():
    # Event Function to show the tk_i image from list
    
    global tkphoto 
    # Canvas Size
    can_h = 400
    can_w = 400
      
    # Get the element i from list
    photo = Image.open(imgs_path[tk_i.get()])
    
    # Get the picture size (widht, height)
    pic_w, pic_h = photo.size

    # Calculate aspect image ratio
    aspect = pic_w/pic_h

    # if picture is wider than taller, the resizing limit will be the picture widht, limited to canvas width
    if aspect > 1:
        res_w = can_w
        res_h = can_w / aspect
    
    # else, the resizing limit will be the picture height, limited to canvas height
    else:
        res_h = can_h
        res_w = can_h * aspect
    

    # resize picture
    photo = photo.resize((int(res_w),int(res_h)))

    # create the Tkinter picture image object
    tkphoto = ImageTk.PhotoImage(photo) 

    # put the picture into the label
    lbl_photo = tk.Label(image=tkphoto,width=can_w,height=can_h,borderwidth=2,relief='solid')
    lbl_photo.grid(row=1,column=4,rowspan=4, padx=10,pady=10)
    
    # display picture number and name
    show_name()
    
    return()

# %%
def show_name():
    # Function to show current image number, total number of selected images and current image name (without path)
    # Path.name method extracts the name from a full 
    img_name = '{} de {}: {}'.format(tk_i.get()+1,len(imgs_path),Path(imgs_path[tk_i.get()]).name)
    lbl_imgname = tk.Label(text=img_name,font=('Consolas 10'))
    lbl_imgname.grid(row=5, column=3,columnspan=3,sticky='NSEW',padx=10,pady=10)
    return()
    

# %%
def go_rgt():
    # Event function to select next (right) image

    if tk_i.get() < len(imgs_path)-1:
        tk_i.set(tk_i.get() + 1)
    else:
        tk_i.set(0)
    
    # Call function to show image number tk_i
    show_imgs()

    return()

# %%
def go_lft():
    # Event function to select previous (left) image

    if tk_i.get() > 0:
        tk_i.set(tk_i.get() - 1)
    else:
        tk_i.set(len(imgs_path)-1)
    
    # Call function to show image number tk_i
    show_imgs()
    
    return()

# %%
def sel_imgs():
    # Event function to select image files
        
    # imgs_path is a global list
    global imgs_path

    images_types = [
            ('Arquivos de imagen','.jpg'),
            ('Arquivos de imagen','.jpeg'),
            ('Arquivos de imagen','.png'),
            ('Arquivos de imagen','.gif'),
            ]

    # Select images
    imgs_path = sorted(list(fd.askopenfilenames(title='Selecione as imagens a enviar',filetypes=images_types)))
    
    # Define show previous image button
    btn_lft = tk.Button(text='<',font=('Consolas 20 bold'),wraplength=100,borderwidth=1,command=go_lft)
    btn_lft.grid(row=1,column=3,rowspan=4,sticky='NSEW',padx=10,pady=10)

    # Define show next image button
    btn_rgt = tk.Button(text='>',font=('Consolas 20 bold'),wraplength=100,borderwidth=1,command=go_rgt)
    btn_rgt.grid(row=1,column=5,rowspan=4,sticky='NSEW',padx=10,pady=10)
        
    # Call function to show first image of selected list
    tk_i.set(0)
    show_imgs()

    return()

# %%
def sel_file():
    # Event function to select Excel file
    
    # contacts_df is the global dataframe with destinataries names, numbers and messages
    global contacts_df

    tk_file_path.set(fd.askopenfilename(
        title='Selecione o arquivo Excel com a lista de destinatarios',
        filetypes=[('Arquivo Excel','.xls'),('Arquivo Excel','.xlsx')]
        )
        )
    
    # Read Excel file
    contacts_df = pd.read_excel(tk_file_path.get(), sheet_name='CLIENTES')

    # Remove rows with empty messages (this improves process ahead)
    contacts_df = contacts_df[~contacts_df['MENSAGEM'].isnull()]

    # Remove rows with empty numbers (this improves process ahead)
    contacts_df = contacts_df[~contacts_df['TELEFONE'].isnull()]

    # Reset index
    contacts_df.reset_index(inplace=True)

    # Keep just the necessary columns
    contacts_df = contacts_df[['CLIENTE','TELEFONE','MENSAGEM']]
    
    # update informations about number of messages to be sent and
    # inform to click button to start process
    lbl_slctdfile['text'] = 'Serão enviadas {} mensagens.\n Clique em "Enviar Mensagens" para iniciar'.format(contacts_df['MENSAGEM'].count())
    
    return()

# %% [markdown]
# Main Process

# %%
def wait_wpp_contacts(timetowaith):
    # Function to wait for WhatsApp contacts side bar for x seconds
    # this indicates that the message text input area is ready to receive messages
    while len(msg_browser.find_elements(By.ID,"side")) < 1:
        tksleep(timetowaith)
    return()

# %%
def stop_sending():
    # Function to stop process
    
    # clean text of "Start Process" button
    btn_send.configure(text='')

    # set flag to stop process    
    tk_stop_sending_messages.set(True)

    return()

# %%
def send_messages():
    # Send messages process
    # basically, this is a Selenium webscripting process, capturing elements from WhatsApp Web
    
    global msg_browser

    # change button label to Stop and activate stop sending function
    btn_send.configure(text='PARAR PROCESSO',command=stop_sending)
    
    # Count total number of messages to send
    msg_total = contacts_df['MENSAGEM'].count()

    # Create intance of Google Chorme browsed
    msg_browser = webdriver.Chrome()
    
    # Navigate to WhatsApp Web
    msg_browser.get("https://web.whatsapp.com/")
    tksleep(2)

    # Link will open the QR Code authorization
    # Wait until user authorization with cell phone
    
    # Wait to load WhatsApp contacts side bar
    # this indicates it is possible to send messages
    wait_wpp_contacts(2)

    init_time = dt.datetime.now()
    lbl_init_time['text'] = 'Tempo de inicio: {}'.format(init_time.strftime('%d-%m-%y %H:%M:%S'))

    fails = 0

    for j, message in enumerate(contacts_df['MENSAGEM']):

        # if stop button was pressed, exit loop
        if tk_stop_sending_messages.get():
            break

        # this version considers all messages are not null
        # dataframe already cleaned up on opening file function

        # Get customer name and number
        name = contacts_df.loc[j,"CLIENTE"]
        phone = contacts_df.loc[j, "TELEFONE"]
        
        # Update status label
        lbl_sending['text'] = 'Enviando mensagem {} de {} ({:.1%})\nPara {} no telefone {}\n{:.1%} de envios falhados'.format(j+1,msg_total,(j+1)/msg_total,name,phone,fails)
        mainwindow.update()

        # Convert message from ASCII into URL plain text
        url_message = urllib.parse.quote(f"{message}")

        # build the link
        link = f"https://web.whatsapp.com/send?phone={phone}&text={url_message}"

        # start tray here?
        try:
            # Get link
            msg_browser.get(link)
            tksleep(3)    

            # check if the sending button is available or the wrong number pops-up
            botao_envio = "span[data-icon='send']"
            botao_telefone_errado = '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button'

            envio_disponivel = False
            telefone_errado = False

            # while sending button is not available OR
            # the wrong number pops-ups
            # stay in this loop
            while (not(envio_disponivel) and not(telefone_errado)):
                
                # capture if sending button is available
                try:
                    msg_browser.find_element(By.CSS_SELECTOR,botao_envio)
                    envio_disponivel = True

                except NoSuchElementException:
                    envio_disponivel = False
                
                # capture if wrong number pop-up
                try:
                    msg_browser.find_element(By.XPATH,botao_telefone_errado)
                    telefone_errado = True
                except NoSuchElementException:
                    telefone_errado = False
                
                tksleep(1)

            # once we have one of the situations (sending button available OR wrong number pop-up)
            # take actions
            # (these options are mutually excludent so never should be both...)
            if envio_disponivel:
                # click on sending button
                msg_browser.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()
                tksleep(2)

                # now start to send images
                for i, img_file in enumerate(imgs_path):
                    # look up for the + sign and click
                    # modificado em 03-Aug-2024
                    # msg_browser.find_element(By.CSS_SELECTOR,"span[data-icon='attach-menu-plus']").click()
                    msg_browser.find_element(By.CSS_SELECTOR,"#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > div > div > span").click() 
                    tksleep(1)
                    
                    # open the image selection window and send the keys
                    msg_browser.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input').send_keys(img_file)
                    tksleep(2)

                    # click on sending button (little triangle)
                    msg_browser.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()
                    tksleep(2)

                # Write on scrolling text box the result of current message sending process
                txt_result = 'Recebeu a mensagem em: {}'.format(dt.datetime.now().strftime('%d-%m-%y %H:%M:%S'))
            
            if telefone_errado:
                # click on ok button on pop-up window to continue
                msg_browser.find_element(By.XPATH,botao_telefone_errado).click()
                txt_result = 'NÃO recebeu a mensagem'
                tksleep(2)
        
        except:
            txt_result = 'NÃO recebeu a mensagem'
                            
        # Print on terminal
        print('{}:{}: {}'.format(j+1,name,txt_result))
            
        # add to result to data frame
        contacts_df.loc[j,'RESULTADO'] = txt_result

        # Write on scrolling text box the result of current message sending process
        txt_sent.insert(tk.INSERT,'{}: {} {}\n'.format(j+1,name,txt_result))
        
        # Point to last line in scrolling text
        txt_sent.see(tk.END)

        # Wait a random time before send next.
        # this is important to avoid WhatsApp to cancel the account due to automation
        tksleep(random.randint(3,7))

        # o loop time medio serve para estimar o eta
        # loop medio = (dt.datetime.now() - init_time) / (j+1)

        # eta deve ser calculada ao final de cada loop
        # eta = tempo agora + loop medio x numero de mensagem que faltam
        eta = dt.datetime.now() + ((dt.datetime.now() - init_time) / (j+1)) * (msg_total-(j+1))

        lbl_eta['text'] = 'Tempo estimado de fim: {}'.format(eta.strftime('%d-%m-%y %H:%M:%S'))

        # get percent of fails until now
        fails = len(contacts_df[contacts_df['RESULTADO'] == 'NÃO recebeu a mensagem'])/(j+1)
    
    # Sending Loop ends here

    # calculate total sent
    total_sent = len(contacts_df[contacts_df['RESULTADO'].str.contains('Recebeu') == True])
    total_fails = len(contacts_df[contacts_df['RESULTADO'] == 'NÃO recebeu a mensagem'])

    # informs that process is finished
    # how many sent, success ratio
    
    lbl_slctdfile['text'] = 'PROCESSO FINALIZADO\n{} Mensagens enviadas\nEnvios falhados: {} ({:.1%})'.format(total_sent,total_fails,fails)
    lbl_sending['text'] = ''
    btn_send.configure(text='')

    # save results dataframe on same location (path) of message file 
    result_file = '{}\Resultado Envios {}.xlsx'.format(Path(tk_file_path.get()).parent,dt.datetime.now().strftime('%d-%m-%y %H-%M-%S'))
    result_df = contacts_df[['CLIENTE','TELEFONE','RESULTADO']]
    result_df.to_excel(result_file,index=False)
    # contacts_df.to_excel(result_file,index=False)
    
    return()

# %% [markdown]
# Main Window Design

# %%
# Create application window
mainwindow = tk.Tk()

# %%
# Main window title
mainwindow.title("Enviar mensagens pelo WhatsApp")

# %%
# Main window label title
lbl_title = tk.Label(text="Enviar mensagens pelo WhatsApp",font=('Consolas 15 bold underline'),borderwidth=1, relief='solid')
lbl_title.grid(row=0, column=0,columnspan=3,sticky='NSEW',padx=10,pady=10)

# %%
# Explaining label
lbl_desc = tk.Label(text=
    """Esta aplicação envia mensagens a través do WhatsApp Web
    junto com imagens, a partir de uma lista em formato Excel.
    A lista deve conter as seguintes colunas:
    NOME, TELEFONE, MENSAGENS, numa folha CLIENTES.
    Cada mensagen pôde ser personalizada. No final, armazena os
    resultados dos envios num outro arquivo Excel na mesma pasta
    do arquivo original."""
    ,font=('Consolas 10'),borderwidth=1, relief='solid')
lbl_desc.grid(row=1, column=0, columnspan=3,sticky='NSEW',padx=10,pady=10) 

# %%
# Excel file selection label
lbl_file = tk.Label(text='Selecione o arquivo Excel com os dados:',font=('Consolas 12'),anchor='e')
lbl_file.grid(row=3,column=0,columnspan=2,sticky='NSEW',padx=10,pady=10)

# %%
# Excel file selecion button
btn_file = tk.Button(text='Clique aqui para selecionar o arquivo',font=('Consolas 10 bold'),wraplength=100,borderwidth=1,command=sel_file)
btn_file.grid(row=3,column=2,sticky='NSEW',padx=10,pady=10)

# %%
# Label with selected Excel file (none at begining, then will show number of message to send)
lbl_slctdfile = tk.Label(text='Sem arquivo selecionado',wraplength=500,font=('Consolas 12'),anchor='center')
lbl_slctdfile.grid(row=4,column=0,columnspan=3,sticky='NSEW',padx=10,pady=10)

# %%
# Image selection label
lbl_imgs = tk.Label(text='Selecione as imagens a enviar:',font=('Consolas 12'),anchor='e')
lbl_imgs.grid(row=2,column=0,columnspan=2,sticky='NSEW',padx=10,pady=10)

# %%
# Image selection button
btn_imgs = tk.Button(text='Clique aqui para selecionar as imagens',font=('Consolas 10 bold'),wraplength=100,borderwidth=1,command=sel_imgs)
btn_imgs.grid(row=2,column=2,sticky='NSEW',padx=10,pady=10)

# %%
# Main process start button
btn_send = tk.Button(text='ENVIAR MENSAGENS',font=('Consolas 10 bold'),command=send_messages)
btn_send.grid(row=5,column=0,columnspan=3,sticky='NSEW',padx=10,pady=10)

# %%
# Current message information status (who, number, total messages)
lbl_sending = tk.Label(text='',wraplength=500,font=('Consolas 12'),anchor='center')
lbl_sending.grid(row=6,column=0,columnspan=3,sticky='NSEW',padx=10,pady=10)

# %%
# Scrolling text to show list of sent messages with success or fail
txt_sent = st.ScrolledText(mainwindow,width = 30, 
                            height = 8, 
                            font = ('Consolas 10'))
txt_sent.grid(row=7,column = 0, columnspan=3,sticky='NSEW', pady = 10, padx = 10)

txt_sent.insert(tk.INSERT,'')
# investigate how to make this read only


# %%
# label for initial time, eta and progress percent
lbl_init_time = tk.Label(text='Tempo inicial',wraplength=500,font=('Consolas 12'),anchor='center')
lbl_init_time.grid(row=6,column=3,columnspan=3, sticky='NSEW',padx=10,pady=10)

lbl_eta = tk.Label(text='Tempo final estimado',wraplength=500,font=('Consolas 12'),anchor='center')
lbl_eta.grid(row=7,column=3,columnspan=3, sticky='NSEW',padx=10,pady=10)

# %%
# 'global' Tkinter IntVar to control current picture to show
tk_i = tk.IntVar(mainwindow,value=0)

# %%
# 'global' Tkinter variable to control stop process
tk_stop_sending_messages = tk.BooleanVar(mainwindow,False)

# %%
# 'global' Tkinter StringVar to store file path
tk_file_path = tk.StringVar(mainwindow,'')

# %%
# define main window icon
mainwindow.iconbitmap(r'icon\whatsapp.ico')

# %%
# Main window
mainwindow.mainloop()


