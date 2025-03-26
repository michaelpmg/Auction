import os
import csv
from docx import Document
from docx.shared import Inches
from os import listdir
from os.path import isfile, join
from datetime import datetime
from pathvalidate import sanitize_filename

def generateDocxForClientSales(client, products, num_fac, result_file_path):
    document = Document()

    date_now = datetime.now();
    date_str = str(date_now.year) + str(date_now.month) + str(date_now.day)
    fac_id_template = "FAC-" + date_str + "-"
    
    # titre du projet
    document.add_heading("FACTURE", level=1).bold = True
    document.add_paragraph('')
    
    # numero facture
    num_fac_p = document.add_paragraph('')
    num_fac_p.add_run('Numero de facture : ' + fac_id_template + str(num_fac)).bold = True
    
    # date
    date_p = document.add_paragraph('')
    date_p_str = str(date_now.year) + '-{:02d}'.format(date_now.month) + '-{:02d}'.format(date_now.day)
    date_p.add_run('date : ' + date_p_str).bold = True

    # client
    date_p = document.add_paragraph('')
    date_p.add_run('Client : ' + client).bold = True

    # details
    document.add_heading("Details de la commande", level=1).bold = True
    document.add_paragraph('')

    fac_total = 0
    
    # list of product
    for product in products:
        prod = products[product]
        prod_qty = prod.qty_
        prod_price = prod.unit_price_
        prod_total = prod_qty * prod_price
        product_str = product + "(x" + str(prod_qty) + " a " + str(prod_price) + "$) = " + str(prod_total) + "$" 
        prod_p = document.add_paragraph('')
        prod_p.add_run(product_str)
        fac_total = fac_total + prod_total

    # total
    total_p = document.add_paragraph('')
    total_p.add_run('Total : ' + str(fac_total) + ".00 $").bold = True

    # skip a line
    document.add_paragraph('')

    # mode paiement
    mode_p = document.add_paragraph('')
    mode_p.add_run('Mode de paiement : ').bold = True
    mode_p.add_run('faire un virement a bestflowers777@gmail.com')

    # password
    pass_p = document.add_paragraph('')
    pass_p.add_run('Mot de passe : ').bold = True
    pass_p.add_run('12345')

    # status
    status_p = document.add_paragraph('')
    status_p.add_run('Statut : ').bold = True
    status_p.add_run('Non payé')

    # skip a line
    document.add_paragraph('')

    # thanks
    thank_p = document.add_paragraph('')
    thank_p.add_run('Merci pour votre achat !').bold = True

    document.save(result_file_path)
    
def getSalesAsDictFromCSV(file_path):
    sales = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            sales.append(row)

    return sales

def ensureDirectoryExists(directory_name):
    os.makedirs(directory_name, exist_ok=True)

class Product:
    def __init__(self):
        self.unit_price_ = 0
        self.qty_ = 0
    
def generateInvoicesFor(cvs_file):
    sales = getSalesAsDictFromCSV(join(to_generate_folder, file))
    
    #get a list of customer from the auction file
    products_by_client = {}
    
    for sale in sales:
        client = sale['Client']
        if client is not "":
            products_by_client[sale['Client']] = {}

    for sale in sales:
        client = sale['Client']
        
        if client == "":
            continue
        
        product_list = {}
        
        #find the client product list or create it
        if products_by_client[client] is not None:
            product_list = products_by_client[client]
        else:
            product_list = {}
            
        #find the current sale product in the product list or create it
        product = Product()
        
        article_name = sale['Article '].strip()
        article_qty = 1;
        article_unit_price = 1;

        try:
            article_qty = int(sale['Quantité'])
        except:
            pass

        try:
            article_unit_price = int(sale['Quantité'])
        except:
            pass
            
        if article_name in product_list:
            product = product_list[article_name]
            product.qty_ = product.qty_ + article_qty
            product.unit_price_ = article_unit_price
        else:
            product = Product()
            product.unit_price_ = article_unit_price
            product.qty_ = article_qty
            
        product_list[article_name] = product
        products_by_client[client] = product_list

    num_client = 0
    print(str(products_by_client))
    for client in products_by_client:
        product_list = products_by_client[client]
        str_now = str(datetime.now().year) + '-{:02d}'.format(datetime.now().month) + '-{:02d}'.format(datetime.now().day)
        result_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Factures_' + str_now )
        ensureDirectoryExists(result_dir)
        result_file = os.path.join(result_dir, sanitize_filename(client + "_" + str_now + ".docx"))
        generateDocxForClientSales(client, product_list, num_client, result_file)
    
def moveEncanFile(csv_file):
    pass
    
if __name__ == "__main__":
    
    to_generate_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Auctions', 'Active')
    files_to_read = [f for f in listdir(to_generate_folder) if isfile(join(to_generate_folder, f)) and f.endswith('.csv')]
    
    for file in files_to_read:
        generateInvoicesFor(file)
