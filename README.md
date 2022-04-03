# NokilonServer
Es un Servidor FTP Creado en VB6, de fácil uso y configuración, con soporte para multiples conexiones, transferencia de archivos mayores de 2GB, soporta conexiones en IPv4 e IPv6. Con la finalidad de ahorrar recursos en la memoria, se ha separado el proyecto en dos partes: Servidor e Interfaz. El programa servidor se ejecuta en segundo plano, la interfaz permite la configuración del servidor (Puerto, Velocidad de transferencia, usuarios) y permite visualizar los comandos enviados por el cliente y las respuestas enviadas por el servidor, tambien permite visualizar las transferencias realizadas y los usuarios conectados. La interfaz se comunica con el programa servidor a traves de mensajes DDE.

Desntro de los plugins, requiere de SQLite3 y J3cnn.dll (Envoltorio personalizado de SQLite3 para VB6).

 ![ITypeComp::Bind](/server/res/nk-server1.jpg)
 
 
 ![ITypeComp::Bind](/server/res/nk-server0.jpg)