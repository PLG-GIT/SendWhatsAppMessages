# Send WhatsApp Web Messages From Excel With Images

This program send messages via WhatsApp Web with images  
The messages must be stored in an Excel file, and mus contain the following columns:  

CLIENTE: Name of destinatary  
TELEFONE: Phone of destinatary  
MENSAGEM: Message to be sent  

Other columns can be present such as name, address, etc, so, by using Excel text concatenation formulae, to send highly personalized messages, including special characters, icons, emoticons, links, etc.  

With messages, the program send the images selected (jpg, png, or gif).  

Notes:  
 - The program works only in Google Chrome  
 - The program waits a random time betweeen messages to avoid WhatsApp to detect automation.  
 - The program displays a scrolling text, showing the historical of messages with sucess or fail.  
 - The program try to send the message and the images, if there is an error, jumps to the next one.  
 - At the end, the program saves an Excel file with same fields (Cliente, Telefone and Mensagem) and adds a column with sucess (and date of message) or fail.  
 - The program is a quite slow, to allow Google Chrome to perform operations.
