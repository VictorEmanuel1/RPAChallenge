from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.chrome import ChromeDriverManager
import time as time
import pandas as pd
import xlrd
import pyautogui as p


def lendo_excel():
    df = pd.read_excel(r'C:\Users\victor.queiroz\Downloads\challenge.xlsx')
    # for i in df.index:
    for i,row in enumerate(df['First Name']):
        first_name = df.loc[i,'First Name']
        last_name = df.loc[i,'Last Name ']
        company_name = df.loc[i,'Company Name']
        role_in_company = df.loc[i,'Role in Company']
        address = df.loc[i,'Address']
        email = df.loc[i,'Email']
        phone_number = df.loc[i,'Phone Number']

        nome = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\fIrst_name.png', confidence =0.85)
        if nome != None:
            print('Nome encotrado')
            p.click(nome.x, nome.y+30)
            p.sleep(0.5)
            p.write(first_name)
   
        else:
            p.scroll(-1000)
            p.sleep(1)
            if nome != None:
                print('Nome encotrado')
                p.click(nome.x, nome.y+30)
                p.sleep(0.5)
                p.write(first_name)
             
            else:
                print('Nome nao encontrado')
            p.scroll(+1000)
        p.sleep(1)
        

        sobrenome = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\last_name.png', confidence =0.85)
        if sobrenome != None:
            print('sobrenome encontrado')
            p.click(sobrenome.x, sobrenome.y+30)
            p.sleep(0.5)
            p.write(last_name)
        else:
            p.scroll(-1000)
            p.sleep(1)
            if sobrenome != None:
                print('sobrenome encontrado')
                p.click(sobrenome.x, sobrenome.y+30)
                p.sleep(0.5)
                p.write(last_name)
            else:
                print('Sobrenome nao encontrado')
            p.scroll(+1000)
        p.sleep(1)

        nome_empresa = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\campany_name.png', confidence = 0.85)
        if nome_empresa != None:
            print('Nome da empresa encontrado')
            p.click(nome_empresa.x, nome_empresa.y+30)
            p.sleep(0.5)
            p.write(company_name)
        else:
            p.scroll(-1000)
            p.sleep(1)
            if nome_empresa != None:
                print('Nome da empresa encontrado')
                p.click(nome_empresa.x, nome_empresa.y+30)
                p.sleep(0.5)
                p.write(company_name)
              
            else:
                print('Nome da empresa nao encontrado')
            p.scroll(+1000)
        p.sleep(1)
        
        
        
        papel_na_empresa = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\role_in_company.png', confidence =0.85)
        if papel_na_empresa != None:
            print('Papel na empresa encontrado')
            p.click(papel_na_empresa.x, papel_na_empresa.y+30)
            p.sleep(0.5)
            p.write(role_in_company)
        else:
            p.scroll(-1000)
            p.sleep(1)
            if papel_na_empresa != None:
                print('Papel na empresa encontrado')
                p.click(papel_na_empresa.x, papel_na_empresa.y+30)
                p.sleep(0.5)
                p.write(role_in_company)
            else:
                print('Papel na empresa nao encontrado')
            p.scroll(+1000)
        p.sleep(1)

        endereco = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\address.png', confidence =0.85)
        if endereco != None:
            print('achei o endereco')
            p.click(endereco.x, endereco.y+30)
            p.sleep(0.5)
            p.write(address)
        else:
            p.scroll(-1000)
            p.sleep(1)
            if endereco != None:
                print('achei o endereco')
                p.click(endereco.x, endereco.y+30)
                p.sleep(0.5)
                p.write(address)
            else:
                print('Nao achei o endereco')
            p.scroll(+1000)
        p.sleep(1)

        email_usuario = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\email.png', confidence =0.85)
        if email_usuario != None:
            print('achei o email do usuario')
            p.click(email_usuario.x, email_usuario.y+30)
            p.sleep(0.5)
            p.write(email)
      
        else:
            p.scroll(-1000)
            p.sleep(1)
            if email_usuario != None:
                print('achei o email do usuario')
                p.click(email_usuario.x, email_usuario.y+30)
                p.sleep(0.5)
                p.write(email)
            else:
                print('Nao achei o email do usuario')
            p.scroll(+1000)
        p.sleep(1)

        telefone = p.locateCenterOnScreen(r'C:\Users\victor.queiroz\Documents\RPAChallenge\Imagem\phone_number.png', confidence = 0.80)
        if telefone != None:
            print('achei o telefone')
            p.click(telefone.x, telefone.y+30)
            p.sleep(0.5)
            p.write(str(phone_number))
        else:
            p.scroll(-1000)
            p.sleep(1)
            if telefone != None:
                print('achei o telefone')
                p.click(telefone.x, telefone.y+30)
                p.sleep(0.5)
                p.write(str(phone_number))
            else:
                print('Nao achei o telefone')
            p.scroll(+1000)
            p.sleep(0.5)
        p.press('enter')
        p.sleep(1)
        

    
        
        
    
    

       

       
def main():
    lendo_excel()


if __name__ == '__main__':
    main()
