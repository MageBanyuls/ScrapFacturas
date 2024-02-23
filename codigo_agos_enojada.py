from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.support.ui import Select
import pandas as pd
import math

archivo_excel=r"C:\Users\USUARIO\Downloads\Empresas a Validar Rentas f.xlsx"


def login_sii (driver,rut,clave):

    
    



        try:
            driver.get("https://zeusr.sii.cl//AUT2000/InicioAutenticacion/IngresoRutClave.html?https://misiir.sii.cl/cgi_misii/siihome.cgi")
    
            ruter_input =  driver.find_element(By.ID, "rutcntr")
            ruter_input.send_keys(rut)
    
            pass_input = driver.find_element(By.ID, "clave")
            pass_input.send_keys(clave)
    
            btn_ingreso = driver.find_element(By.ID, "bt_ingresar")
            btn_ingreso.click()

            time.sleep(2)
    
            try:
                alert = driver.switch_to.alert
                alert.dismiss()
            except:
                pass
    
            try:
                alert = driver.switch_to.alert
                alert.dismiss()
            except:
                pass
    
            try:
                driver.find_element(By.ID, "titulo")
                print("no se pudo hacer login")
                return False
    
            except NoSuchElementException:
                print('login exitosos')
                #aca analizo si esta la pantalla de siguiente"
                try:
                    try:
                        modal = driver.find_element(By.CSS_SELECTOR, 'div.modal-dialog')
    
                        if modal:
                            # Si hay un modal, hacer clic en el botón de cierre
                            btn_cierre_modal = driver.find_element(By.XPATH, '//*[@id="ModalEmergente"]/div/div/div[3]/button')
                            btn_cierre_modal.click()
                    except:
                        pass
    
                    time.sleep(2)
    
                    try:
                        modal = driver.find_element(By.ID,'myMainCorreoVigente')
                        if modal.is_displayed():
                            driver.execute_script("arguments[0].style.display = 'none';", modal)
    
                    except:
                        pass
    
    
                except NoSuchElementException:
                    boton_siguiente = driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/p[2]/a[1]')
                    boton_siguiente.click()
    
                    try:
                        try:
                            alert = driver.switch_to.alert
                            alert.dismiss()
                        except:
                            pass
    
                        try:
                            alert = driver.switch_to.alert
                            alert.dismiss()
                        except:
                            pass
    
                        try:
                            modal = driver.find_element(By.CSS_SELECTOR, 'div.modal-dialog')
    
                            if modal:
                                # Si hay un modal, hacer clic en el botón de cierre
                                btn_cierre_modal = driver.find_element(By.XPATH, '//*[@id="ModalEmergente"]/div/div/div[3]/button')
                                btn_cierre_modal.click()
                        except:
                            pass
    
                        time.sleep(2)
    
                        try:
                            modal = driver.find_element(By.ID,'myMainCorreoVigente')
                            if modal.is_displayed():
                                driver.execute_script("arguments[0].style.display = 'none';", modal)
                        except:
                            pass
    
                    except NoSuchElementException:
                        print('keseste erroooor!!')
                return True
    
        except TimeoutException as timeout_error:
    
            print(f"Error de tiempo de espera para RUT {rut}: {timeout_error}")
        except Exception as e:
            print(f"Error durante el proceso para RUT {rut}: {e}")


def getData(parte_df):

    options = Options()
    
    options.add_argument("--headless")
    driver_path = r"C:\Users\USUARIO\Downloads\chromedriver-win64\chromedriver.exe"
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    for index, row in parte_df.iterrows():
        rut=row['Rut']
        clave=row['Clave SII']

        login_success=login_sii(driver,rut,clave)
        time.sleep(1)
    
        #busqueda facturas de compra/venta
    
        driver.get("https://www4.sii.cl/consdcvinternetui/#/index")
    
        #configurar el seelct del mes
        select_mes = driver.find_element(By.ID,"periodoMes")
    
        time.sleep(1)
        select1 = Select(select_mes)
        #select1.select_by_value("11")
    
        #select del anio
        select_anio = driver.find_element(By.XPATH,'//*[@id="my-wrapper"]/div[2]/div[1]/div[1]/div/div[1]/div/div[3]/div/form/div[2]/select[2]')
        time.sleep(1)
        select2 = Select(select_anio)
        #select2.select_by_value("2023")
        select_rut = driver.find_element(By.XPATH,'//*[@id="my-wrapper"]/div[2]/div[1]/div[1]/div/div[1]/div/div[3]/div/form/div[1]/select')
        time.sleep(1)
        
        btn_siguiente = driver.find_element(By.XPATH,'//*[@id="my-wrapper"]/div[2]/div[1]/div[1]/div/div[1]/div/div[3]/div/form/div[3]/button')
    
    
        btn_seccion_venta = driver.find_element(By.CSS_SELECTOR,'a[ui-sref="venta"]')
    
        btn_seccion_compra = driver.find_element(By.CSS_SELECTOR,'a[ui-sref="compra"]')
    
        #Periodos de tiempo buscados
        periodos = [{"anio":"2023","mes":"11"},{"anio":"2023","mes":"12"},{"anio":"2024","mes":"01"}]
    
        total_facturas_de_venta = 0
        total_facturas_de_compra = 0
    
        for periodo in periodos:
            #
            time.sleep(0.5)
            select1.select_by_value(periodo["mes"])
            time.sleep(0.5)
            select2.select_by_value(periodo["anio"])
            select_rut.send_keys(rut)
            
    
            btn_siguiente.click()
            time.sleep(2)
    
            try:
                table_body = driver.find_element(By.XPATH,'//*[@id="home"]/table/tbody[2]')
                time.sleep(1)
                rows = table_body.find_elements(By.CSS_SELECTOR,'tr.ng-scope')
                print('mes ' + periodo["mes"])
                cantidad = 0
                for row in rows:
                    #print('fila')
                    tipo = row.find_element(By.CSS_SELECTOR,'td[scope="row"]').text
                    cols = row.find_elements(By.CSS_SELECTOR,'td.ng-binding')
                    col = cols[0].text
                    valor = int(col)
                    if "(39)" in tipo or "(48)" in tipo or "(35)" in tipo or "(38)" in tipo or "(41)" in tipo:
                        print("no hay que sumarlo")
                    else:
                        cantidad = cantidad + valor
    
                    #cantidad = cantidad + valor
                    print({"tipo":tipo,"valor":col,"mes":periodo["mes"],"anio":periodo["anio"]})
    
                print('canntidad final facturas de compra')
                print(cantidad)
                periodo["facturas_de_compra"] = cantidad
                total_facturas_de_compra = total_facturas_de_compra + cantidad
    
            except NoSuchElementException:
                print('no hay nada')
                periodo["facturas_de_compra"] = 0
    
    
            btn_seccion_venta.click()
            time.sleep(3)
            #accion de venta
            try:
                #
                table_body = driver.find_element(By.XPATH,'//*[@id="home"]/table/tbody[2]')
                time.sleep(3)
                rows = table_body.find_elements(By.CSS_SELECTOR,'tr.ng-scope')
                print('mes '+periodo["mes"])
                cantidad = 0
                for row in rows:
                    #print('fila')
                    tipo = row.find_element(By.CSS_SELECTOR,'td[scope="row"]').text
                    cols = row.find_elements(By.CSS_SELECTOR,'td.ng-binding')
                    col = cols[0].text
                    valor = int(col)
                    if "(39)" in tipo or "(48)" in tipo or "(35)" in tipo or "(38)" in tipo or "(41)" in tipo:
                        print("no hay que sumarlo")
                    else:
                        cantidad = cantidad + valor
                    #cantidad = cantidad + valor
                    print({"tipo":tipo,"valor":col,"mes":periodo["mes"],"anio":periodo["anio"]})
    
                print('canntidad final facturas de venta')
                print(cantidad)
                #resultados.append({"facturasVenta":cantidad})
                periodo["facturas_de_venta"] = cantidad
                total_facturas_de_venta = total_facturas_de_venta + cantidad
    
    
            except NoSuchElementException:
                print('no hay nada')
                periodo["facturas_de_venta"] = 0
    
            time.sleep(3)
            btn_seccion_compra.click()
    
    
    
        print(periodos)
    
    
        promedio_facturas_de_compra = total_facturas_de_compra/3
        promedio_facturas_de_compra_redondeado = math.ceil(promedio_facturas_de_compra)
    
        promedio_facturas_de_venta = total_facturas_de_venta/3
        promedio_facturas_de_venta_redondeado = math.ceil(promedio_facturas_de_venta)
    
    
    
        print(f"facturas de compra: {total_facturas_de_compra}")
        print(f"promedio facturas de compra: {promedio_facturas_de_compra} / redondeado {promedio_facturas_de_compra_redondeado}")
    
        print(f"facturas de venta: {total_facturas_de_venta}")
        print(f"promedio facturas de venta: {promedio_facturas_de_venta} / redondeado {promedio_facturas_de_venta_redondeado}")
    
    
        time.sleep(2)
    
    
        #logica para recopilar
    
        #Boletas de Honorarios Electrónicas recibidas(INFORMES DE BOLETAS RECIBIDAS)
        driver.get("https://loa.sii.cl/cgi_IMT/TMBCOC_MenuConsultasContribRec.cgi?dummy=1461943244650")
        select_anio = driver.find_element(By.XPATH,'/html/body/div[2]/center/table[3]/tbody/tr[2]/td[2]/div/font/select')
        time.sleep(1)
        select = Select(select_anio)
        select.select_by_value("2023")
        btn = driver.find_element(By.ID,'cmdconsultar124')
        time.sleep(1)
        btn.click()
        promedio_boletas_anuales_recibidas = 0
        promedio_boletas_anuales_recibidas_redondeado = 0
        try:
            boletas_anuales_txt = driver.find_element(By.XPATH,'/html/body/div[3]/center/table[2]/tbody/tr[6]/td/table/tbody/tr[15]/td[2]/font').text
            boletas_anuales = int(boletas_anuales_txt)
            promedio_boletas_anuales_recibidas = boletas_anuales/12
            promedio_boletas_anuales_recibidas_redondeado = math.ceil(promedio_boletas_anuales_recibidas)
            
    
        except NoSuchElementException:
            print('no hay')
            
    
    
        driver.get("https://zeus.sii.cl/cvc_cgi/bte/bte_indiv_cons?1")
 
        select_anio = driver.find_element(By.ID,'ANOA')
        time.sleep(1)
        select = Select(select_anio)
        select.select_by_value("2023")
    
        btn = driver.find_element(By.XPATH,'/html/body/center[2]/form/table/tbody/tr[2]/td[3]/font/input[1]')
        time.sleep(1)
        btn.click()
        promedio_boletas_anuales_emitidas = 0
        promedio_boletas_anuales_emitidas_redondeado = 0
        try:
            boletas_anuales_txt = driver.find_element(By.XPATH,'/html/body/center[2]/form[1]/table/tbody/tr[15]/td[4]/div/font').text
            boletas_anuales = int(boletas_anuales_txt)
            promedio_boletas_anuales_emitidas = boletas_anuales/12
            promedio_boletas_anuales_emitidas_redondeado = math.ceil(promedio_boletas_anuales_emitidas)
    
        except NoSuchElementException:
            print('no hay')
    
        time.sleep(2)
        df.at[index,'Promedio Facturas Venta']=promedio_facturas_de_venta_redondeado
        df.at[index,'Promedio Facturas Compra']=promedio_facturas_de_compra_redondeado
        df.at[index,'Promedio de boletas recibidas']=promedio_boletas_anuales_recibidas_redondeado
        df.at[index,'Promedio boletas emitidas']=promedio_boletas_anuales_emitidas_redondeado
        if login_success:
            df.at[index,'Revisado']='Si'
        else:
            df.at[index,'Revisado']='Contraseña'

        df.to_excel(archivo_excel,index=False)
    driver.quit()

        
        
 
 
 
 

        
        
    
 
 


 
 
 
 
if __name__ == "__main__":
 
    df=pd.read_excel(archivo_excel)

    df_no = df[df['Revisado'] == 'No']
    # Dividir el DataFrame en partes para distribuir entre los hilos
    num_threads = 5
    particiones = np.array_split(df_no, num_threads)


    # Configurar el número máximo de hilos
    max_workers = 5

    # Crear un ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=max_workers) as executor:

        futures = []

        # Lanzar una tarea para cada porción del DataFrame
        for parte_df in particiones:
            future = executor.submit(getData, parte_df)
            futures.append(future)

        # Esperar a que todas las tareas se completen
        for future in futures:
            future.result()