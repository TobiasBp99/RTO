import webbrowser
import pandas as pd
import lxml
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver import ActionChains
from datetime import datetime
from datetime import timedelta
import requests
from urllib.request import urlopen
import time


def main():
    l_inicio = []
    l_resultado = []
    l_fin = []
    l_tipo = []
    l_uso = []
    l_categoria = []
    l_planilla = []
    l_certificado = []
    l_vence = []
    l_dominio = []
    l_res = []
    l_cuit = []
    l_razon = []
    l_linea = []
    l_fecha = []
    l_trabajo = []

    l_head = []
    flag_exc_hea = 0
    i = 0
    xpath_header    = "/html/body/section[2]/div/div[2]/div/span/fieldset/div[4]/table/thead/tr/th[1]"

    flag_exc_row = 0
    flag_exc_col = 0
    j = 0
    k = 0
    xpath_table     = "/html/body/section[2]/div/div[2]/div/span/fieldset/div[4]/table/tbody/tr[1]/td[1]"
    
    # datos login

    user = GetUsername()
    passwd = GetPassword()
    # datos de la pagina a entrar
    url = "https://rto.cent.gov.ar/rto/RTO/planillaDiaria"
    headers = {"User-Agent":"Mozilla/5.0"}

    browser = webdriver.Firefox(executable_path=r"./geckodriver")
    browser.get(url)

    response = requests.get(url,headers = headers)
    print(response.status_code)
    
    # Me logueo
    Login(browser,user,passwd)
    
    # Ya estoy en las planillas
    # Selecciono que no se agrupen
    ShowSheet(browser)
    # Pido datos de comienzo del control
    day_from,month_from,year_from = GetDateFrom()
    s_from = FormatDate(day_from,month_from,year_from)
    day_to, month_to, year_to = GetDateTo()
    s_to = FormatDate(day_to, month_to, year_to)
    
    fecha_from = datetime.strptime(s_from, '%d/%m/%Y')
    fecha_to = datetime.strptime(s_to, '%d/%m/%Y')
    #--------------------------------------------------------------#
    SetDate(browser, fecha_from)
    #Leo head, creo que no sirve para nada
    while flag_exc_hea == 0:
        try:
            s_head = browser.find_element_by_xpath(xpath_header.replace("/th["+'1'+']',"/th["+str(i+1)+']')).text
            l_head.append(s_head)
            i += 1
        except:
            flag_exc_hea = 1
    
        
    
    while fecha_from != fecha_to:
        flag_exc_col = 0
        flag_exc_row = 0
        while flag_exc_row == 0:        #Leo row
            while flag_exc_col == 0:    #Leo col
                try:                    #Para ver si termine la fila
                    s_col = browser.find_element_by_xpath(xpath_table.replace('/tr['+'1'+']'+'/td['+'1'+']', '/tr['+str(k+1)+']'+'/td['+str(j+1)+']')).text
                    
                    if j == 0:
                        l_fecha.append(fecha_from)
                        l_inicio.append(s_col)
                    elif j == 1:
                        l_resultado.append(s_col)
                    elif j == 2:
                        l_fin.append(s_col)
                        l_trabajo.append( str(datetime.strptime(l_fin[k],"%H:%M") - datetime.strptime(l_inicio[k],"%H:%M")))
                    elif j == 3:
                        l_tipo.append(s_col)
                    elif j == 4:
                        l_uso.append(s_col)
                    elif j == 5:
                        l_categoria.append(s_col)
                    elif j == 6:
                        l_planilla.append(s_col)
                    elif j == 7:
                        l_certificado.append(s_col)
                    elif j == 8:
                        l_vence.append(s_col)
                    elif j == 9:
                        l_dominio.append(s_col)
                    elif j == 10:
                        l_res.append(s_col)
                    elif j == 11:
                        l_cuit.append(s_col)
                    elif j == 12:
                        l_razon.append(s_col)
                    elif j == 13:
                        l_linea.append(s_col)
                    j += 1
                except:                 #Fin de fila
                    flag_exc_col = 1
                    j = 0

            k += 1                      #Leo la siguente fila
            flag_exc_col = 0            #Limpio flag
                
            try:                        #para ver si termine las columnas
                s_next = browser.find_element_by_xpath(xpath_table.replace('/tr['+'1'+']'+'/td['+'1'+']', '/tr['+str(k+1)+']'+'/td['+str(j+1)+']')).text
                """
                l_day.append(l_car)     #agrego el individuo
                while (len(l_car)):     #Limpio el individuo
                    l_car.pop()
                """
            except:                     #Fin de columnas
                flag_exc_row = 1
                k = 0
        
        fecha_from = fecha_from + timedelta(days=1)
        SetDate(browser, fecha_from)
    #Salgo del while

    t_work = datetime.strptime("00:00","%H:%M")
    q_m1 = 0
    q_m2 = 0
    q_m3 = 0
    q_n1 = 0
    q_n2 = 0
    q_n3 = 0
    q_o2 = 0
    q_o3 = 0
    q_o4 = 0
    q_up = 0
    q_apto = 0
    q_cond = 0
    q_rec = 0 
    q_anul = 0
    q_empty = 0


    for x in range(len(l_res)):
        res = l_res[x]
        t = l_trabajo[x]
        c = l_categoria[x]

        if res == "Apto":
            q_apto +=1
        elif res == "Condicional":
            q_cond +=1
        elif res == "Rechazado":
            q_rec +=1
        elif res == "Anulado":
            q_anul +=1

    
        s_t = str(t)
        s_min = s_t[2:4]
        s_sec = s_t[5:]
        t_work += timedelta(minutes=int(s_min),seconds=int(s_sec))
        
        if c == "M1":
            q_m1 +=1
        elif c == "M2":
            q_m2 +=1
        elif c == "M3":
            q_m3 +=1
        elif c == "N1":
            q_n1 +=1
        elif c == "N2":
            q_n2 +=1
        elif c == "N3":
            q_n3 +=1
        elif c == "O2":
            q_o2 +=1
        elif c == "O3":
            q_o3 +=1
        elif c == "O4":
            q_o4 +=1
        elif c == "UP":
            q_up +=1
        else:
            q_empty +=1
    
        
    
    data = {"Fecha":l_fecha,"Tiempo Trabajado":l_trabajo,"Inicio":l_inicio,"Resultados":l_resultado,"Fin":l_fin,"Tipo":l_tipo,"Uso":l_uso,"Categoria":l_categoria,"Planilla":l_planilla,"Certificado":l_certificado,"Vencimiento":l_vence,"Dominio":l_dominio,"Resultado":l_res,"CUIT":l_cuit,"Razón Social":l_razon,"Línea":l_linea}
    df = pd.DataFrame(data)
    

    cont = {"M1":[q_m1],"M2":[q_m2],"M3":[q_m3],"N1":[q_n1],"N2":[q_n2],"N3":[q_n3],"O2":[q_o2],"O3":[q_o3],"O4":[q_o4],"UP":[q_up],"-":[q_empty]}
    cf = pd.DataFrame(cont)

    s_twork = str(t_work)
    
    s_th = s_twork[11:13]
    s_tm = s_twork[14:16]
    s_ts = s_twork[17:]
    
    
   
    qtotal = len(l_res)
    t_work += timedelta(minutes=int(s_min),seconds=int(s_sec))
    avg = (timedelta(hours=int(s_th),minutes=int(s_tm),seconds=int(s_ts)))/qtotal
    #avg = 0 #hardcode

    stat = {"Tiempo Trabajado":[s_th+':'+s_tm+':'+s_ts],"Promedio":[str(avg)],"Revisiones":[len(l_res)]}
    sf = pd.DataFrame(stat)

    
    
    file_name = GetFileName()
    xlwritter = pd.ExcelWriter(file_name+".xlsx")    # pylint: disable=abstract-class-instantiated
    df.to_excel(xlwritter,sheet_name = 'Tabla',index = False)
    cf.to_excel(xlwritter,sheet_name = 'Cantidad',index = False)
    sf.to_excel(xlwritter,sheet_name = 'Tiempo',index = False)
    xlwritter.save()
    xlwritter.close()

    browser.close()
        
        




def GetPassword():
    return(input("Password\t:"))

def GetUsername():
    return(input("Username\t:"))

def GetDateFrom():
    print("\nInicio")
    d = int(input("Dia\t\t:"))
    m = int(input("Mes\t\t:"))
    y = int(input("año\t\t:"))
    return( d , m , y)

def GetDateTo():
    print("\nFin")
    d = int(input("Dia\t\t:"))
    m = int(input("Mes\t\t:"))
    y = int(input("año\t\t:"))
    return( d , m , y)

def FormatDate(d,m,y):
    fecha = ''
    if d <10 :
        fecha += '0'
    fecha += str(d)
    fecha += '/'

    if m < 10:
        fecha += '0'
    fecha += str(m)
    fecha += '/'

    fecha+= str(y)
    return (fecha)

def SetDate(browser,fecha_from):
    elegir_f = browser.find_element_by_name("fechaConsulta")
    elegir_f.click()
    elegir_f.clear()
    elegir_f.send_keys(datetime.strftime(fecha_from, '%d/%m/%Y'))
    elegir_f.send_keys(Keys.ENTER)
    browser.find_element_by_id("box_f").click()


def Login(browser,user,passwd):
    username = browser.find_element_by_name("j_username")
    password = browser.find_element_by_name("j_password")
    username.send_keys(user)
    password.send_keys(passwd)
    browser.find_element_by_id('submit').click()

def ShowSheet(browser):
    agrupar_por_linea = Select(browser.find_element_by_name("null"))
    agrupar_por_linea.select_by_visible_text("NO")

def GetFileName():
    return(input("File name\t:"))







main()