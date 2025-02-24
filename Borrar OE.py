import win32com.client

# Crear una instancia de SAP GUI Scripting
SapGuiAuto = win32com.client.GetObject("SAPGUI")

# Obtener la instancia de la aplicación SAP
application = SapGuiAuto.GetScriptingEngine

# Conectar al sistema SAP especificado en la ruta
connection = application.OpenConnection("Ingresa tu conexión", True)

# Obtener la sesión activa
session = connection.Children(0)

# Ingresar usuario y pass 
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "Ingresar usuario"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Ingresa tu contraseña"
# Enter
session.findById("wnd[0]").sendVKey(0)



#Ingresar las OE a borrar
OE = [

0000000000,
1111111111,
2222222222,


]

for x in OE:
    try:
        #Seleccionar VL02N
        #session.findById("wnd[0]").resizeWorkingPane(143, 41, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        # Enter
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = x
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        
        
        print(x,",", session.findById("wnd[0]/sbar").text)
    except:  
        next
        print(x,",", session.findById("wnd[0]/sbar").text) 

#Cerrar sesion SAP
session.findById("wnd[0]").close()
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
