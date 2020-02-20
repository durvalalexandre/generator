import time

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

PRODUCTION = False

#USERNAME = "P0642979"
#PASSWORD = "tp3itn6t"
USERNAME = "default"
PASSWORD = "password"
FILTER = ""
TEST_SERVER = "nt2670:8080"
PRODUCTION_SERVER = "gengax"
ITEM_LIST = [

'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_051',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_052',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_053',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_054',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_055',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_056',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_057',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_058',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_059',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_060',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_061',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_062',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_063',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_064',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_065',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_107',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_108',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_109',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_119',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_120',
'OD_COR_POTENC_SUCURSAL_RIO_BRANCO_122'

]

driver = webdriver.Chrome()
driver.maximize_window()
actions = ActionChains(driver)

if PRODUCTION:
    server = PRODUCTION_SERVER
else:
    server = TEST_SERVER

driver.get(f"http://{server}/gax/#/!/list:opm.paramgroups")
assert "Genesys" in driver.title

while True:
    try:
        username_field = driver.find_element_by_name("username")
        username_field.send_keys(USERNAME)
        break
    except Exception as e:
        # print(e)
        time.sleep(1)
        continue

password_field = driver.find_element_by_name("password")
password_field.send_keys(PASSWORD)
submit_button = driver.find_element_by_css_selector("button.g-buttonbar-item-first")
submit_button.click()
time.sleep(6)
search_field = None

while True:
    try:
        search_field = driver.find_element_by_name("quickfilter")
        break
    except NoSuchElementException:
        time.sleep(2)
        continue

# Do actions for every item in list
done = []
not_done = []
time.sleep(8)
for item in ITEM_LIST:
    search_field.clear()
    time.sleep(1)
    search_field.send_keys(item)
    time.sleep(1)
    c = 0

    try:
        panel = driver.find_element_by_class_name('slick-viewport')
        panel.click()
        actions.send_keys(Keys.PAGE_DOWN, Keys.PAGE_DOWN, Keys.PAGE_DOWN).perform()
        elements = driver.find_elements_by_xpath("//div[contains(@class, 'slick-cell l1 r1')]")
        for i, el in enumerate(elements):
            if el.text == item:
                el.click()
                break
    except NoSuchElementException as e:
        time.sleep(1)
        continue
    except StaleElementReferenceException as e:
        time.sleep(1)
        continue
    c += 1
    if c > 4:
        print('Could not continue automatization.')
        exit(0)

    time.sleep(7.5)

    try:
        elements = driver.find_elements_by_css_selector("input[type='radio'][value='transaction']")
        if len(elements) > 0:
            for el in elements:
                parent = el.find_element_by_xpath('./..')
                parent.click()
                time.sleep(0.05)
        button_span = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[4]/div[3]/div[2]/div/div[3]/button[1]")
        button = button_span.find_element_by_xpath('./..')
        button.click()
        time.sleep(20)
        done.append(item)
    except NoSuchElementException as e:
        print(e.msg)
        print(f'Erro ao executar o item "{item}". Verifique as opções no GA e tente novamente.')
        print('Geralmente este erro ocorre por algum problema nas options que foram carregadas nesta lista no GA.')
        print('Talvez um arquivo de áudio com nome ao invés de número em homologação?')
        not_done.append(item)

print('Completados:')
[print(item) for item in done]
print(f'Total: {len(done)} itens.')
print('\nErro:')
[print(item) for item in not_done]
print(f'Total: {len(not_done)} itens.')
print('\nFIM DA EXECUÇÃO.')