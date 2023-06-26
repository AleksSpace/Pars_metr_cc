import xlsxwriter
from selenium import webdriver
import time

from selenium.common import NoSuchElementException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService


def check_exists_element(web_driver, by, xpath):
    try:
        web_driver.find_element(by=by, value=xpath)
    except NoSuchElementException:
        return False
    return True


def get_target_data(web_driver, url_product):
    # Переходим по ссылкам на товары и собираем нужные данные
    id_product = ''
    name_product = ''
    regular_price = ''
    promo_price = ''
    brand_product = ''
    url_product = url_product
    # Если у товара есть промо цена, то записываем её и регулярную цену
    if check_exists_element(web_driver, By.XPATH, '//div[@class="product-prices-block__top"]'):
        regular_price_product = web_driver.find_element(
            by=By.XPATH, value='//div[@class="product-prices-block__top"]//span[@class="product-price__sum"]/span'
        )
        regular_price = f'{regular_price_product.text} р'

        promo_price_product = web_driver.find_element(
            by=By.XPATH,
            value='//div[@class="product-price-discount-above__bottom"]//span[@class="product-price__sum"]/span'
        )
        promo_price = f'{promo_price_product.text} р'

        brand = web_driver.find_element(by=By.XPATH,
                                        value='//span[@class="product-attributes__list-item-value"]')
        brand_product = brand.text

    else:
        regular_price_product = web_driver.find_element(
            by=By.XPATH,
            value='//div[@class="product-price-discount-above__bottom"]//span[@class="product-price__sum-rubles"]'
        )
        regular_price = f'{regular_price_product.text} р'

        # Бренд
        brand = web_driver.find_element(
            by=By.XPATH, value='//li[@class="product-attributes__list-item"]/a')
        brand_product = brand.text

    # Название продукта
    name = web_driver.find_element(by=By.XPATH, value='//div[@class="page-subcategory__wrapper"]/div[2]//h1/span')
    name_product = name.text

    # ID продукта
    article_number = web_driver.find_element(by=By.XPATH,
                                             value='//div[@class="page-subcategory__wrapper"]/div[2]//div/p')
    id_product = article_number.text.split(': ')[1]

    data_list = [id_product, name_product, url_product, regular_price, promo_price, brand_product]

    return data_list


def pars_metro(web_driver, target_url, target_city):
    try:
        # Заходим на страницу выбранной категории продуктов в Metro
        web_driver.get(url=target_url)
        time.sleep(5)

        # В появившемся окне нажимаем "Смотреть каталог"
        button_catalog = web_driver.find_element(by=By.CLASS_NAME, value='shop-select-dialog__item')
        button_catalog.find_element(by=By.TAG_NAME, value='button').click()
        time.sleep(3)

        # В появившемся окне нажимаем кнопку "Выбрать магазин"
        button_shop = web_driver.find_element(by=By.XPATH, value='//*[@id="__layout"]/div/div/div[7]/div[2]/div[4]/button[2]')
        button_shop.click()
        time.sleep(3)

        # В поле выбора города подставиться город из списка
        select = web_driver.find_element(by=By.CLASS_NAME, value='multiselect__select')
        select.click()
        time.sleep(3)
        multiselect = web_driver.find_element(by=By.CLASS_NAME, value='multiselect__tags')
        input_city = multiselect.find_element(by=By.TAG_NAME, value='input')
        input_city.send_keys(target_city)
        input_city.send_keys(Keys.ENTER)

        # нажимаем кнопку Сохранить
        web_driver.find_element(by=By.XPATH,
                                value='//div[@class="pickup__apply-btn-desk"]/button').click()
        time.sleep(5)

        paginate = web_driver.find_elements(by=By.XPATH, value='//nav[@class="subcategory-or-type__pagination"]/ul/li')

        # Нажимаем кнопку "Показать ещё" пока не отобразятся все товары
        for _ in range(len(paginate) - 2):
            time.sleep(2)
            web_driver.find_element(by=By.XPATH, value='//*[@id="catalog-wrapper"]/main/div[3]/button').click()
            time.sleep(2)

        # Собираем все ссылки на товары и записываем их в список
        url_target_products = []
        products = web_driver.find_elements(by=By.XPATH, value='//div[@class="subcategory-or-type__products"]/div')
        for product in products:
            product_card = product.find_element(by=By.CLASS_NAME, value='product-card-photo__content')
            product_url = product_card.find_element(by=By.TAG_NAME, value='a').get_attribute('href')
            url_target_products.append(product_url)

        list_for_xls = [
            ['ID продукта', 'Название продукта', 'Ссылка на товар', 'Регулярная цена', 'Промо цена', 'Бренд']
        ]

        for url_target_product in url_target_products[::-1]:
            web_driver.get(url=url_target_product)
            time.sleep(5)
            # Если товара нет в наличии, то переходим к следующему
            if check_exists_element(web_driver, By.XPATH, '//div[@class="product-page-content__prices-block"]/p'):
                continue
            else:
                list_for_xls.append(get_target_data(web_driver, url_target_product))

        return list_for_xls

    except Exception as ex:
        print(ex)
    finally:
        web_driver.close()
        web_driver.quit()


if __name__ == "__main__":
    url = 'https://online.metro-cc.ru/category/zamorozhennye-produkty/zamorozhennye-gotovye-blyuda'
    cities = ['Москва', 'Санкт-Петербург']

    with xlsxwriter.Workbook('data_pars_metro.xlsx') as workbook:
        worksheet_msk = workbook.add_worksheet(cities[0])
        worksheet_spb = workbook.add_worksheet(cities[1])
        for city in cities:
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            the_path = 'chromedriver/chromedriver.exe'
            service = ChromeService(executable_path=the_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # Записываем данные в файл
            if city == 'Москва':
                list_product_msk = pars_metro(driver, url, city)
                for row_num, data in enumerate(list_product_msk):
                    worksheet_msk.write_row(row_num, 0, data)
            else:
                list_product_spb = pars_metro(driver, url, city)
                for row_num, data in enumerate(list_product_spb):
                    worksheet_spb.write_row(row_num, 0, data)
