# Required packages
import openpyxl
import pandas as pd
import sys

# Very primitive mapping for column names to their index value
dataFields = {
    'feed_product_type': 0,
    'item_sku': 1,
    'brand_name': 2,
    'external_product_id': 3,
    'item_name': 4,
    'manufacturer': 5,
    'recommended_browse_nodes': 6,
    'standard_price': 7,
    'main_image_url': 8,
    'other_image_url1': 9,
    'other_image_url2': 10,
    'other_image_url3': 11,
    'other_image_url4': 12,
    'other_image_url5': 13,
    'other_image_url6': 14,
    'other_image_url7': 15,
    'other_image_url8': 16,
    'swatch_image_url': 17,
    'parent_child': 18,
    'parent_sku': 19,
    'style': 20,
    'description_style': 21,
    'rooms': 22,
    'part_number': 23,
    'product_description': 24,
    'short_product_description': 25,
    'color_name': 26,
    'color_map': 27,
    'pattern_name': 28,
    'size_name': 29,
    'material_type': 30,
    'catalog_number': 31,
    'generic_keywords': 32,
    'bullet_point1': 33,
    'bullet_point2': 34,
    'bullet_point3': 35,
    'bullet_point4': 36,
    'bullet_point5': 37,
    'style_name': 38,
    'unit_count': 39,
    'unit_count_type': 40,
    'country_of_origin': 41,
    'currency': 42,
    'condition_type': 43,
    'merchant_shipping_group_name': 44,
}

# HTML strings for replacing/appending
main_html = '<!DOCTYPE html><html lang=de><meta charset=UTF-8><meta content="width=device-width,initial-scale=1" name=viewport><link rel=preconnect href=https://fonts.gstatic.com><link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;500;600;700;800;900&display=swap" rel=stylesheet><link href="https://fonts.googleapis.com/css2?family=Hind:wght@300;500;700&display=swap" rel=stylesheet><link rel=stylesheet type=text/css href=style.css><title>Bricoflor</title><div class=Bricoflor-body><div class=Bricoflor-header><div class=Bricoflor-headerInnerContent><div class=Bricoflor-headerTextWrapper><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/tickk.png alt="Tick icon"> <span class=Bricoflor-headerText>Große Auswahl mit über 30,000 Produkten</span></div><div class=Bricoflor-headerTextWrapper><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/tickk.png alt="Tick icon"> <span class=Bricoflor-headerText>Fachkompetenz seit Jahrenzehnten</span></div><div class=Bricoflor-headerTextWrapper><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/tickk.png alt="Tick icon"> <span class=Bricoflor-headerText>Bestpreisgarantie</span></div></div></div><div class=Bricoflor-InnerContent><div class=Bricoflor-logoSection><div class=Bricoflor-leftLogoSection><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/CompanyLogo.png alt="Company logo" style="width: 100%"></div><div class=Bricoflor-rightLogoSection><div class=Bricoflor-LargeScreenLargeScreenShortContact><div class=Bricoflor-LargeScreenCallerFaceLogo><div class=Bricoflor-LargeScreenCallerFaceLogo><img class="" src=caller.png alt="Caller icon"></div><div><div class=Bricoflor-LargeScreenCaller>Fachberater:</div><div class=Bricoflor-LargeScreenCallerNumber>0202 69508170</div><div class=LargeScreenTiming>Mo.-So.: 8 - 20 Uhr</div></div></div></div><div class=Bricoflor-searchIcon><!-- <i class="fas fa-search"></i> --></div></div></div><div class=Bricoflor-shortContact><div class=Bricoflor-Callernumber><div class=Bricoflor-callerFaceLogo><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/caller.png alt="Caller icon" style="width: 100%"></div><div class=Bricoflor-caller>Fachberater:</div><div class=Bricoflor-callerNumber>0202 69508170</div></div><div class=Bricoflor-timing>Mo.-So.: 8 - 20 Uhr</div></div><div class=Bricoflor-ImgTagsDescripWrapper><img class=Bricoflor-mainImage src={{main_image_url}} alt="Main product image"><div class=Bricoflor-tagsAndDescriptionWrapper><div class=Bricoflor-ProductDescriptionWrapper><div class=Bricoflor-longTitle>{{item_name}}</div><ul class=Bricoflor-DescriptionPoints><li class=Bricoflor-DescriptionOneWrapper><div class=Bricoflor-leaf></div><div class=Bricoflor-DescriptionOne>{{main_bullet_one}}</div></li><li class=Bricoflor-DescriptionOneWrapper><div class=Bricoflor-leaf></div><div class=Bricoflor-DescriptionOne>{{main_bullet_two}}</div></li><li class=Bricoflor-DescriptionOneWrapper><div class=Bricoflor-leaf></div><div class=Bricoflor-DescriptionOne>{{main_bullet_three}}</div></li><li class=Bricoflor-DescriptionOneWrapper><div class=Bricoflor-leaf></div><div class=Bricoflor-DescriptionOne>{{main_bullet_four}}</div></li></ul></div></div></div><div class=Bricoflor-detailsWrapper><div class=Bricoflor-detailsDescription><strong>{{description_style}}</strong></div><div class=Bricoflor-detailsDescription>Um das Beste aus Ihren Räumen rauszuholen und einen stimmigen Look zu erhalten, kombinieren Sie zu Ihrer ausgewählten Tapete die passenden Unis oder weitere vorgeschlagene Mustervariationen.</div><div class=Bricoflor-detailsDescription>Sie sind nicht wegzudenken, wenn es um die Gestaltung von Wänden geht. Tapeten sind neben dem Streichen die Hauptmöglichkeit, eine Wand zu gestalten und ihr ein neues Leben zu verleihen.</div></div></div></div><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><!-- {{product}} --><div class=Bricoflor-helpWrapper><div class=Bricoflor-helpText><div class=Bricoflor-helpTitle>Wir sind persönlich für Sie da!</div><div class=Bricoflor-helpDescription>Unser kompetentes BRICOFLOR-Team berät Sie gerne bei Ihrem Projekt und bietet Ihnen Unterstützung bei Ihrem Einkauf. Für einen reibungslosen Ablauf Ihres Projekts bis in den Einkauf.</div><div class=Bricoflor-helpContact><div class=Bricoflor-helpIcon><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/telephone.png alt="Telephone icon" style="width: 100%"></div><div><div class=Bricoflor-helpNumber>0202-69508170</div><div class=Bricoflor-helpTiming>Mo.-So.: 8-20 Uhr</div></div></div><div class=Bricoflor-helpContact><div class=Bricoflor-helpIcon><img class="" src=https://www.bricoflor.de/media/wysiwyg/ebay2021/letter.png alt="Letter icon" style="width: 100%"></div><div><div class=Bricoflor-helpNumber>info@bricoflor.de</div><div class=Bricoflor-helpTiming>E-mail</div></div></div></div><div class=Bricoflor-helpImg><img class=Bricoflor-ManImg src=https://www.bricoflor.de/media/wysiwyg/ebay2021/man.png alt="Man holding flooring image"></div></div><div class=Bricoflor-servicesWrapper><div class=Bricoflor-BottomLogoSection><img class=Bricoflor-BottomLogoImg src=https://www.bricoflor.de/media/wysiwyg/ebay2021/CompanyLogo.png alt="Girl in a jacket"></div><div class=Bricoflor-deliveryBrandsSection><div class=Bricoflor-deliveryBrandTitle>Versand</div><div class=Bricoflor-brandsWrapper><div class=Bricoflor-brand><img class=Bricoflor-brandImg src=https://www.bricoflor.de/media/wysiwyg/ebay2021/ups.PNG alt="UPS Logo"></div><div class=Bricoflor-brand><img class=Bricoflor-brandImg src=https://www.bricoflor.de/media/wysiwyg/ebay2021/dhl.PNG alt="DHL Logo"></div></div></div></div><div class=Bricoflor-footer><div class=Bricoflor-footerInnerContent><div class=Bricoflor-footerLongText>All Preis inkl. gesetzt. Mehrwertsteuer zzgl. Versandkosten und ggf. Nachnahnegebühren, wenn nicht anders beschrieben.</div><div class=Bricoflor-footerShortText>© 2021 - BRICOFLOR GmbH</div></div></div></html>'
product_html = '<div class="Bricoflor-technicalSpecs"> <div class="Bricoflor-technicalSpecsInner"> <img class="Bricoflor-SampleImg1" src="{{product_image}}" alt="Product image"/> <div class="Bricoflor-technicalSpecsTitleWrapper"> <div class="Bricoflor-technicalSpecsShortTitle">{{short_product_description}}</div><div class="Bricoflor-technicalSpecsTitle">{{title}}</div><div class="Bricoflor-specsWrapper"> <div class="Bricoflor-specTitle"> Stil: </div><div class="Bricoflor-line"></div><div class="Bricoflor-specDetail">{{style}}</div></div><div class="Bricoflor-specsWrapper"> <div class="Bricoflor-specTitle"> Größe: </div><div class="Bricoflor-line"></div><div class="Bricoflor-specDetail">{{size_name}}</div></div><div class="Bricoflor-specsWrapper"> <div class="Bricoflor-specTitle"> Farbe: </div><div class="Bricoflor-line"></div><div class="Bricoflor-specDetail">{{color_map}}</div></div><div class="Bricoflor-specsWrapper"> <div class="Bricoflor-specTitle"> Ideal für Räume: </div><div class="Bricoflor-line"></div><div class="Bricoflor-specDetail">{{bullet_point3}}</div></div></div></div><div class="Bricoflor-priceDeliveryWrapper"> <div class="Bricoflor-priceWrapper"> <div class="Bricoflor-priceTagIcon"> <img class="" src="https://www.bricoflor.de/media/wysiwyg/ebay2021/tag.png" alt="Tags icon" style="width: 100%"/> </div><div class="Bricoflor-priceDescription"> <div class="Bricoflor-totalPrice"> <span class="Bricoflor-price">{{standard_price}}€</span>pro Tapete </div><div class="Bricoflor-individualPrice"> entspricht <strong>{{unit_price}}€</strong> pro m² </div></div></div><div class="Bricoflor-priceDeliveryLine"></div><div class="Bricoflor-deliveryWrapper"> <div class="Bricoflor-wagonIcon"> <img class="" src="https://www.bricoflor.de/media/wysiwyg/ebay2021/van1.JPG" alt="Van icon" style="width: 100%"/> </div><div class="Bricoflor-deliveryDescription"> <div class="Bricoflor-deliveryPrice"> <strong>Versand:</strong> Kostenlos (außer Inseln) </div><div class="Bricoflor-deliveryTime"> <strong>LIeferzeit:</strong> 3-5 Arbeitstage </div></div></div></div></div>'

# Data variable / file import
export_data = pd.DataFrame([])
import_file = pd.read_csv('data.csv', sep=';')
import_rows_processed = 0
export_rows_count = 0

# Function for appending product details
def append_item_details(item, row_details):
    item = item.replace("{{short_product_description}}", str(row_details[dataFields['short_product_description']]), 1)
    item = item.replace("{{title}}", get_external_id_with_curly_braces(str(row_details[dataFields['external_product_id']])), 1)
    item = item.replace("{{product_image}}", str(row_details[dataFields['main_image_url']]), 1)
    item = item.replace("{{style}}", str(row_details[dataFields['style']]).rstrip(','), 1)
    item = item.replace("{{color_map}}", str(row_details[dataFields['color_map']]), 1)
    item = item.replace("{{size_name}}", str(row_details[dataFields['size_name']]), 1)
    item = item.replace("{{bullet_point3}}", str(row_details[dataFields['rooms']]), 1)
    item = item.replace("{{standard_price}}", get_external_id_with_curly_braces(str(row_details[dataFields['external_product_id']]), 'Price'), 1)
    item = item.replace("{{unit_price}}", get_external_id_with_curly_braces(str(row_details[dataFields['external_product_id']]), 'PricePerSquare'), 1)
    return item

def assemble_children_string(array):
    return " ".join(str(child) + ',' for child in array)

def get_external_id_with_curly_braces(e_product_id, append_string=None):
    if append_string:
        return "{{" + append_string + e_product_id + "}}"
    else:
        return "{{" + e_product_id + "}}"

print("*** Script Starting ***")

# Row loop through the CSV file
for i, row in import_file.iterrows():
    print("Beginning import file data loop...")
    import_rows_processed = import_rows_processed + 1

    # Skip the parent divider rows
    if row[dataFields['parent_child']] != "parent":
        current_parent = row[dataFields['parent_sku']]
        item_sku = row[dataFields['item_sku']]
        external_id = row[dataFields['external_product_id']]
        # Copy of the main HTML string
        item_html = main_html

        # Replace dynamic details for the main product
        item_html = item_html.replace("{{item_name}}", get_external_id_with_curly_braces(str(row[dataFields['external_product_id']])), 1)
        item_html = item_html.replace("{{description_style}}", str(row[dataFields['description_style']]), 1)
        item_html = item_html.replace("{{main_image_url}}", str(row[dataFields['main_image_url']]), 1)

        item_html = item_html.replace("{{main_bullet_one}}", str(row[dataFields['style']]).rstrip(','), 1)
        item_html = item_html.replace("{{main_bullet_two}}", str(row[dataFields['size_name']]), 1)
        item_html = item_html.replace("{{main_bullet_three}}", str(row[dataFields['color_name']]), 1)
        item_html = item_html.replace("{{main_bullet_four}}", str(row[dataFields['rooms']]), 1)

        # While loop states to allow loop breaking, repeats until true returned
        backward_finished = False
        forward_finished = False
        count = i
        children = []

        # Loop through products the previous products for any related products until we hit the parent product
        while not backward_finished:
            current_row = import_file.loc[count - 1]
            i_product_html = product_html

            if current_row[dataFields['parent_child']] == "parent":
                # Skip, parent product row
                backward_finished = True
                break
            else:
                # Related products found, HTML replace and append
                i_product_html = append_item_details(i_product_html, current_row)
                item_html = item_html.replace("<!-- {{product}} -->", str(i_product_html), 1)
                children.append(str(current_row[dataFields['external_product_id']]))
                count = count - 1

        # Reset count ready to go backwards
        count = i

        # Loop through products the next products for any related products until we hit the parent product
        while not forward_finished:
            if count + 1 > len(import_file) - 1:
                break

            current_row = import_file.loc[count + 1]
            i_product_html = product_html

            if current_row[dataFields['parent_child']] == "parent":
                # Skip, parent product row
                forward_finished = True
                break
            else:
                # Related products found, HTML replace and append
                i_product_html = append_item_details(i_product_html, current_row)
                item_html = item_html.replace("<!-- {{product}} -->", str(i_product_html), 1)
                children.append(str(current_row[dataFields['external_product_id']]))
                count = count + 1

        # Append edited HTML to Excel file row
        export_data = export_data.append(pd.DataFrame({'Item SKU': item_sku, 'External ID': external_id, 'Related': assemble_children_string(children), 'HTML': item_html}, index=[0]), ignore_index=True)
        export_rows_count = export_rows_count + 1
        print("HTML Generated for item (" + item_sku + ")")

# Save to Excel file
export_data.to_excel('export_data.xlsx')

print("*** Script Ended ***")
print("*** Import rows processed: " + str(import_rows_processed) + ' Export rows added to export file: ' + str(export_rows_count) + ' ***')