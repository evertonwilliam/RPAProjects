#import rarfile

#dir = 'C:\RPATemp\Python\Projetos_RPA_Python\Py_04\exe\MundoRPA - Desafio App.rar'
#ext = 'C:\RPATemp\Python\Projetos_RPA_Python\Py_04\exe'
#try:
#    r = rarfile.RarFile(dir)
#    r.extractall(ext)
#    r.close()
#except Exception as ex:
#    print(ex)
    
#backend='win32'
#backend='uia'
 
 
from pywinauto.application import Application
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard

exe = r'C:\RPATemp\Python\Projetos_RPA_Python\Py_04\exe\MundoRPA - Desafio App.exe'

app = Application().start(exe).connect(title='Mundo RPA', timeout=1)

app.MundoRpa.wait('ready')

app.MundoRpa.Button0.click()

#texEditor1 = app.MundoRpa.child_window(title='ID', auto_id="4", control_type='tkchild').wrapper_object()
#texEditor1.type_keys("teste")

#win = app.MundoRpa
#win.wait('ready')

#button = win.Button0.click()
#button.click()

#app.edit1.set_text("Example-utf8.txt")

app.kill()



