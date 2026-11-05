"""Articles transcoding tool (AS400 -> D365 data migration templates)."""

import numpy as np
import pandas as pd
import gradio as gr


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _norm_cols(df):
    df.columns = df.columns.str.strip().str.upper()
    return df


def _clean(s):
    """AS400 CHAR fields are space-padded: strip + NaN -> ''."""
    return s.fillna('').astype(str).str.strip()


def _lookup(keys, ref, key_col, val_col):
    ref = ref.copy()
    ref[key_col] = _clean(ref[key_col])
    mapping = ref.drop_duplicates(subset=key_col).set_index(key_col)[val_col]
    return _clean(keys).map(mapping).fillna('')


def _flag(s, mapping):
    return _clean(s).map(mapping).fillna('')


def _fmt_date(s):
    """20230627 -> '27/06/2023 00:00' ; empty or invalid -> ''"""
    d = pd.to_datetime(_clean(s), format='%Y%m%d', errors='coerce')
    return d.dt.strftime('%d/%m/%Y 00:00').fillna('')


# ---------------------------------------------------------------------------
# DM PRODUCTS
# ---------------------------------------------------------------------------

def process_products_file(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        output_data = pd.DataFrame({
            'PRODUCTNUMBER': '',
            'ISPRODUCTKIT': 'NO',
            'PRODUCTDESCRIPTION': products_data['B9BQT1'].replace("  ", ""),
            'PRODUCTNAME': products_data['B9BQT1'].replace("  ", ""),
            'PRODUCTSEARCHNAME': '',
            'PRODUCTSUBTYPE': 'PRODUCT',
            'PRODUCTTYPE': 'ITEM',
            'PRODUCTAS400NUMBER': products_data['B9N6N2'],
        })
        output_file = 'PRODUIT.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


# ---------------------------------------------------------------------------
# DM PRODUCT TRANSLATION
# ---------------------------------------------------------------------------

def process_products_translation(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        replacements = {
            'F': 'FR',
            'A': 'En-Us',
            'B': 'AR',
            'C': 'zh-hans',
            'D': 'DE',
            'E': 'ES',
            'P': 'pt-PT',
            'R': 'RU',
            'Z': 'FR',
        }
        products_data['B9CACD'] = products_data['B9CACD'].str.strip().replace(replacements, regex=False)

        output_data = pd.DataFrame({
            'PRODUCTNUMBER': '',
            'LANGUAGEID': products_data['B9CACD'],
            'DESCRIPTION': products_data['B9BRT1'],
            'PRODUCTNAME': products_data['B9BPT1'],
            'PRODUCTLAGNUMBER': products_data['B9N6N2'],
        })
        output_file = 'PRODUIT_TRANSLATION.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


# ---------------------------------------------------------------------------
# DM CATEGORY ASSIGNMENT
# ---------------------------------------------------------------------------

def process_products_category_assignment(products_file):
    try:
        products_data = pd.read_excel(products_file, dtype=str)
        products_data.columns = products_data.columns.str.strip().str.upper()

        categorie = products_data['CLASSE(S) ATC'].fillna('').str.strip()

        output_data = pd.DataFrame({
            'PRODUCTNUMBER': '',
            'PRODUCTCATEGORYNAME': categorie,
            'PRODUCTCATEGORYHIERARCHYNAME': np.where(categorie != '', 'Hierarchy ATC', ''),
            'PRODUCTNUMBER1': products_data['B9N6N2'],
        })
        output_file = 'PRODUIT_CATEGORY_ASSIGNMENT.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {e}") from e


# ---------------------------------------------------------------------------
# DM EXTERNAL ITEM DESCRIPTIONS
# ---------------------------------------------------------------------------

def process_products_external_items_for_vendors(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        output_data = pd.DataFrame({
            'PRODUCTNUMBER': '',
            'VENDORACCOUNTNUMBER': products_data['C3L7NB'],
            'VENDORPRODUCTDESCRIPTION': '',
            'VENDORPRODUCTNUMBER': products_data['C3NNN2'],
        })
        output_file = 'PRODUIT_External_item_descriptions_for_vendors.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


def process_products_external_items_for_customers(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        output_data = pd.DataFrame({
            'ITEMNUMBER': '',
            'CUSTOMERACCOUNTNUMBER': products_data['C3L7NB'],
            'CUSTOMERPRODUCTDESCRIPTION': '',
            'CUSTOMERPRODUCTNUMBER': products_data['C3NNN2'],
        })
        output_file = 'PRODUIT_External_item_descriptions_for_customers.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


# ---------------------------------------------------------------------------
# DM RELEASED PRODUCTS
# ---------------------------------------------------------------------------

def process_released_products(products_file, rembss_file):
    try:
        products_data = _norm_cols(pd.read_excel(products_file, dtype=str))
        rembss_data = _norm_cols(pd.read_excel(rembss_file, dtype=str, sheet_name="remboursement"))
        code_douane_data = _norm_cols(pd.read_excel(rembss_file, dtype=str, sheet_name="code_douane"))
        transco_pays = _norm_cols(pd.read_excel(rembss_file, dtype=str, sheet_name="PaysOrigine"))
        code_pospro_data = _norm_cols(pd.read_excel(rembss_file, dtype=str, sheet_name="PositionProduit"))

        def get_code_rembss():
            try:
                return _lookup(products_data['B9VIC1'], rembss_data, 'EKVIC1', 'EKIEPC')
            except Exception as e:
                raise Exception(f"Erreur get_code_rembss: {e}") from e

        def get_code_douane():
            try:
                return _lookup(products_data['B9W0CD'], code_douane_data, 'J5W0CD', 'J5W1CD')
            except Exception as e:
                raise Exception(f"Erreur get_code_douane: {e}") from e

        def get_code_iso():
            try:
                return _lookup(products_data['B9YFC2'], transco_pays,
                               'CODEVALEURSOURCEAS400', 'CODEVALEURCIBLED365')
            except Exception as e:
                raise Exception(f"Erreur get_code_iso: {e}") from e

        def get_code_provenance():
            try:
                return _lookup(products_data['B9YGC2'], transco_pays,
                               'CODEVALEURSOURCEAS400', 'CODEVALEURCIBLED365')
            except Exception as e:
                raise Exception(f"Erreur get_code_provenance: {e}") from e

        def get_pospro():
            try:
                return _lookup(products_data['B9BUCD'], code_pospro_data, 'AGBUCD', 'AGANTX')
            except Exception as e:
                raise Exception(f"Erreur get_pospro: {e}") from e

        amm = _clean(products_data['B9JOS2'])
        hmd = _clean(products_data['B9JVS2'])
        rfitem = _clean(products_data['B9IZS2'])
        marque = _clean(products_data['B9JST2'])

        output_data = pd.DataFrame({
            'HSOSOCIALSECURITYRATE': get_code_rembss(),
            'HSOBRANDEDGENERIC': _flag(products_data['B9JSS2'], {'P': 'PR', 'G': 'G', '': 'G'}),
            'HSOGALENICFORM': products_data['GALENICFORM'],
            'HSOPRESCRIPTION': np.where(_clean(products_data['B9JXS2']) != '', 'YES', 'NO'),
            'HSOPSYCHOQTY': products_data['B9KGQT'],
            'HSOINSPECTEDITEM': '',
            'HSOINSPECTIONCOMPANY': '',
            'HSOINSPECTIONTYPE': '',
            'ITEMNUMBER': '',
            'PRODUCTSEARCHNAME': '',
            'HSOSTORAGETEMP': '',
            'HSOSUFFIX': '',
            'INVENTORYUNITSYMBOL': 'PCS',
            'ORIGINCOUNTRYREGIONID': get_code_iso(),
            'PRIMARYVENDORACCOUNTNUMBER': _clean(products_data['NUMSUPP']),
            'PRODUCTGROUPID': '',
            'PRODUCTVOLUME': (pd.to_numeric(products_data['B9KTQT'], errors='coerce')
                              * pd.to_numeric(products_data['B9KVQT'], errors='coerce')
                              * pd.to_numeric(products_data['B9KUQT'], errors='coerce')).fillna(0),
            'PURCHASEUNITSYMBOL': '',
            'SALESPRICE': '',
            'SALESPRICEDATE': '',
            'UNITCOST': '',
            'UNITCOSTDATE': '',
            'GROSSDEPTH': products_data['B9KTQT'],
            'HSOAMMSTART': '',
            'HSODOPINGITEM': '',
            'GROSSPRODUCTHEIGHT': products_data['B9KVQT'],
            'GROSSPRODUCTWIDTH': products_data['B9KUQT'],
            'HSOLPPEND': '',
            'HSOLPPRATE': '',
            'HSOLPPRCODE': '',
            'HSOLPPSTART': '',
            'INTRASTATCOMMODITYCODE': '',
            'NETPRODUCTWEIGHT': products_data['B9KXQT'],
            'HSOOFFICIALGENERIC': _flag(products_data['B9JIS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'PRODUCTLIFECYCLESTATEID': get_pospro(),
            'HSOPHARMACATEGORY': '',
            'HSOAMM': np.where(amm == 'A', 'YES', 'NO'),
            'HSOEXIPIENT': _flag(products_data['B9JPS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'HSOSERVICETYPE': '',
            'HSOHUMANVETERINARY': _flag(products_data['B9IYS2'], {'H': 'Humain', 'V': 'Vétérinaire'}),
            'HSOMEDICALDEVICE': np.where(hmd == 'M', 'YES', 'NO'),
            'HSOTRADEPRICESTART': '',
            'HSONARCOTIC': np.where(hmd == 'S', 'YES', 'NO'),
            'HSOVIGILANCE': '',
            'ALTERNATIVEITEMNUMBER': '',
            'ALTERNATIVEPRODUCTUSAGECONDITION': 'RIEN EN STOCK',
            'APPROVEDVENDORCHECKMETHOD': '',
            'ARETRANSPORTATIONMANAGEMENTPROCESSESENABLED': 'YES',
            'ARRIVALHANDLINGTIME': '',
            'BASEPRICEPURCHASE': '',
            'BASESALESPRICESOURCE': '',
            'BATCHMERGEDATECALCULATIONMETHOD': '',
            'BATCHNUMBERGROUPCODE': '',
            'BESTBEFOREPERIODDAYS': '',
            'BUYERGROUPID': '',
            'CARRYINGCOSTABCCODE': '',
            'COSTCALCULATIONGROUPID': '',
            'COSTCHARGESQUANTITY': '',
            'COSTGROUPID': '',
            'CUSTOMSDESCRIPTION': '',
            'DEFAULTDIRECTDELIVERYWAREHOUSE': '',
            'DEFAULTLEDGERDIMENSIONDISPLAYVALUE': '',
            'DEFAULTORDERTYPE': 'PURCH',
            'DEFAULTRECEIVINGQUANTITY': '',
            'EIGHTHPRODUCTFILTERCODE': '',
            'FIFTHPRODUCTFILTERCODE': '',
            'FIRSTPRODUCTFILTERCODE': '',
            'FIRSTPRODUCTFILTERGROUPID': '',
            'FIXEDCOSTCHARGES': '',
            'FIXEDPURCHASEPRICECHARGES': '',
            'FOURTHPRODUCTFILTERCODE': '',
            'FREIGHTALLOCATIONGROUPID': '',
            'INVENTORYCOUNTINGREASONCODEPOLICYNAME': '',
            'INVENTORYRESERVATIONHIERARCHYNAME': '',
            'ISBATCHNUMBERACTIVE': 'YES',
            'ISDELIVEREDDIRECTLY': '',
            'HSOOFFICIALJOURNALDATE': _fmt_date(products_data['B9CHDX']),
            'ISINSTALLMENTELIGIBLE': '',
            'ISINTERCOMPANYPURCHASEUSAGEBLOCKED': 'NO',
            'ISINTERCOMPANYSALESUSAGEBLOCKED': 'NO',
            'ISMANUALDISCOUNTPOSREGISTRATIONPROHIBITED': '',
            'ISRESTRICTEDFORCOUPONS': '',
            'ISSALESPRICEADJUSTMENTALLOWED': '',
            'ISSALESPRICEINCLUDINGCHARGES': '',
            'ISSALESWITHHOLDINGTAXCALCULATED': '',
            'ISSERIALNUMBERACTIVE': '',
            'ISSHIPALONEENABLED': '',
            'ISUNITCOSTAUTOMATICALLYUPDATED': 'YES',
            'ISUNITCOSTINCLUDINGCHARGES': '',
            'ITEMMODELGROUPID': 'STOCK',
            'ITEMOVERUNDERDELIVERYTOLERANCEGROUPID': '',
            'LANDEDCOSTTYPEGROUPID': '',
            'LOWERWARRANTABLEPRICERANGELIMIT': '',
            'MARGINABCCODE': '',
            'MAXIMUMPICKQUANTITY': '',
            'NINTHPRODUCTFILTERCODE': '',
            'PRODUCTCOVERAGEGROUPID': '',
            'PRODUCTDIMENSIONGROUPNAME': '',
            'PRODUCTNUMBER': '',
            'PRODUCTSUBTYPE': 'PRODUCT',
            'PRODUCTTYPE': 'ITEM',
            'PURCHASECHARGEPRODUCTGROUPID': '',
            'PURCHASECHARGESQUANTITY': '',
            'PURCHASEITEMWITHHOLDINGTAXGROUPCODE': '',
            'PURCHASELINEDISCOUNTPRODUCTGROUPCODE': '',
            'PURCHASEMULTILINEDISCOUNTPRODUCTGROUPCODE': '',
            'PURCHASEOVERDELIVERYPERCENTAGE': '',
            'PURCHASEPRICE': '',
            'PURCHASEPRICEDATE': '',
            'PURCHASEPRICEQUANTITY': '0',
            'PURCHASEPRICETOLERANCEGROUPID': '',
            'PURCHASEPRICINGPRECISION': '',
            'PURCHASEREBATEPRODUCTGROUPID': '',
            'PURCHASESALESTAXITEMGROUPCODE': '',
            'PURCHASESUPPLEMENTARYPRODUCTPRODUCTGROUPID': '',
            'PURCHASEUNDERDELIVERYPERCENTAGE': '',
            'REVENUEABCCODE': '',
            'REVRECBUNDLE': '',
            'REVRECDEFAULTREVENUERECOGNITIONSCHEDULE': '',
            'REVRECEXCLUDEFROMCARVEOUT': '',
            'REVRECMEDIANPRICE': '',
            'REVRECMEDIANPRICEMAXIMUMTOLERANCE': '',
            'REVRECMEDIANPRICEMINIMUMTOLERANCE': '',
            'REVRECREVENUERECOGNITIONENABLED': '',
            'REVRECREVENUETYPE': '',
            'SALESCHARGEPRODUCTGROUPID': '',
            'SALESCHARGESQUANTITY': '',
            'SALESITEMWITHHOLDINGTAXGROUPCODE': '',
            'SALESLINEDISCOUNTPRODUCTGROUPCODE': '',
            'SALESMULTILINEDISCOUNTPRODUCTGROUPCODE': '',
            'SALESOVERDELIVERYPERCENTAGE': '',
            'SALESPRICECALCULATIONCHARGESPERCENTAGE': '',
            'SALESPRICECALCULATIONCONTRIBUTIONRATIO': '',
            'SALESPRICECALCULATIONMODEL': '',
            'SALESPRICEQUANTITY': '',
            'SALESPRICINGPRECISION': '',
            'SALESREBATEPRODUCTGROUPID': '',
            'SALESSALESTAXITEMGROUPCODE': '',
            'SALESSUPPLEMENTARYPRODUCTPRODUCTGROUPID': '',
            'SALESUNDERDELIVERYPERCENTAGE': '',
            'SALESUNITSYMBOL': '',
            'SCALEINDICATOR': '',
            'SEARCHNAME': '',
            'SECONDPRODUCTFILTERCODE': _flag(products_data['B9JKS2'], {'O': 'FROID', 'N': 'Ambiant'}),
            'SECONDPRODUCTFILTERGROUPID': '',
            'SELLENDDATE': '',
            'SELLSTARTDATE': '',
            'SERIALNUMBERGROUPCODE': '',
            'SERVICETYPE': '',
            'SEVENTHPRODUCTFILTERCODE': '',
            'SHELFADVICEPERIODDAYS': '',
            'HSOPSYCHOTROPIC': np.where(hmd == 'Q', 'YES', 'NO'),
            'SIXTHPRODUCTFILTERCODE': '',
            'STORAGEDIMENSIONGROUPNAME': 'WMS',
            'TAREPRODUCTWEIGHT': '',
            'TAXRATETYPE': '',
            'TENTHPRODUCTFILTERCODE': '',
            'THIRDPRODUCTFILTERCODE': '',
            'TRACKINGDIMENSIONGROUPNAME': 'BATCH',
            'TRANSFERORDERLANDEDCOSTGROUPID': '',
            'TRANSFERORDEROVERDELIVERYPERCENTAGE': '10',
            'TRANSFERORDERUNDERDELIVERYPERCENTAGE': '10',
            'UNITCONVERSIONSEQUENCEGROUPID': '',
            'UNITCOSTQUANTITY': '',
            'VALUEABCCODE': '',
            'VENDORINVOICELINEMATCHINGPOLICY': '',
            'VOYAGEARRIVALGROUPID': '',
            'WAREHOUSEBATCHDETAILSCAPTUREPOLICY': '',
            'WAREHOUSEFILTERCODE': '',
            'WAREHOUSEFILTERGROUP': '',
            'WAREHOUSEMOBILEDEVICEDESCRIPTIONLINE1': '',
            'WAREHOUSEMOBILEDEVICEDESCRIPTIONLINE2': '',
            'WARRANTABLEPRICERANGEBASETYPE': '',
            'WARRANTYDURATIONTIME': '',
            'WARRANTYDURATIONTIMEUNIT': '',
            'WILLPICKINGWORKBENCHAPPLYBOXINGLOGIC': '',
            'WILLTOTALPURCHASEDISCOUNTCALCULATIONINCLUDEPRODUCT': '',
            'WILLTOTALSALESDISCOUNTCALCULATIONINCLUDEPRODUCT': '',
            'HSOREEFERITEM': np.where(rfitem == '3', 'YES', 'NO'),
            'HSOADMINDOSE': '',
            'HSOAI': '',
            'HSOAMMEND': '',
            'HSOBATCHREFUNDABLE': '',
            'HSOSUBJECTTFR': _flag(products_data['B9KCS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'HSOVACCINE': _flag(products_data['B9JLS2'], {'O': 'YES', 'N': 'NO'}),
            'HSOBUYER': '',
            'HSOCATALOGITEM': '',
            'HSOCEMARK': '',
            'HSOCOEF': '',
            'HSOCUSTOMSRATE': '',
            'HSODISTRIBUTOR': '',
            'HSODPAV': '',
            'ISHAZARDOUSMATERIAL': _flag(products_data['B9JJS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'SHELFLIFEPERIODDAYS': _clean(products_data['B9O1N2']),
            'HSOHYBRIDDRUG': '',
            'HSOICEPACK': _flag(products_data['B9JKS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'HSOICEPACKDURATION': '',
            'HSOACTIVEINGREDIENT': '',
            'HSOBLOODDERIVED': np.where(hmd == 'G', 'YES', 'NO'),
            'HSOEXCEPTIONALDRUG': _flag(products_data['B9KAS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'HSOITEMACTIVITY': '',
            'HSOICPECODE': '',
            'HSOITEMDLVMODE': _flag(products_data['B9B0CD'], {'A': 'AIR', 'M': 'SEA', '': 'SEA'}),
            'HSOITEMRANGE': '',
            'HSOITEMSOURCE': get_code_provenance(),
            'HSOLABORATORYID': '',
            'HSOLEGACYPRODUCTNUMBER': _clean(products_data['B9N6N2']),
            'HSOLIFECYCLELASTCHANGE': '',
            'HSOLIFECYCLENEXTCHANGE': '',
            'HSOMAD': '',
            'HSOINITHOSPITAL': '',
            'HSOMEDICALSURGICAL': '',
            'HSONSDSCODE': '0',
            'HSOOLDLPPR': '',
            'HSOISCYTOTOXIC': np.where(hmd == 'C', 'YES', 'NO'),
            'HSOITEMBRAND': np.where(marque == 'M', 'MARQUE', ''),
            'HSOPLVTMET': '',
            'HSOPLVTTER': '',
            'HSOMEDDEVCLASS': '',
            'HSOORPHANDRUG': _flag(products_data['B9JMS2'], {'O': 'YES', 'N': 'NO', '': 'NO'}),
            'HSOPHARMARAWMAT': np.where(hmd == 'Y', 'YES', 'NO'),
            'HSOPRODUCTTYPE': _clean(products_data['TYPEPRODUIT']),
            'HSOPRECURSOR': np.where(hmd == 'Z', 'YES', 'NO'),
            'HSOREGNATURE': '',
            'HSOREPLACINGITEM': '',
            'HSOPRINCEPSGENERIC': _flag(products_data['B9JSS2'], {'P': '1', 'G': '2', '': '0'}),
            'HSOTRADEPRICEEND': '',
            'HSOTRANSPORTKO': '',
            'HSOTRANSPORTTEMP': '',
            'HSOABCPHARMACEUTICAL': '',
            'HSOWHOLESALERCATID': '',
        })

        output_file = 'PRODUIT_Released.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file

    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {e}") from e


def transform_released_products_create(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        output_data = pd.DataFrame({
            'ITEMNUMBER': '',
            'BOMUNITSYMBOL': 'PCS',
            'INVENTORYRESERVATIONHIERARCHYNAME': 'BATCH',
            'INVENTORYUNITSYMBOL': 'PCS',
            'ISPRODUCTKIT': 'NO',
            'ITEMMODELGROUPID': 'STOCK',
            'LOWERWARRANTABLEPRICERANGELIMIT': '',
            'PRODUCTDESCRIPTION': products_data['B9BQT1'].str.strip(),
            'PRODUCTDIMENSIONGROUPNAME': '',
            'PRODUCTGROUPID': '',
            'PRODUCTNAME': products_data['B9BQT1'].str.strip(),
            'PRODUCTNUMBER': '',
            'PRODUCTSEARCHNAME': '',
            'PRODUCTSUBTYPE': 'PRODUCT',
            'PRODUCTTYPE': 'ITEM',
            'PURCHASESALESTAXITEMGROUPCODE': '',
            'PURCHASEUNITSYMBOL': '',
            'RETAILPRODUCTCATEGORYNAME': '',
            'SALESSALESTAXITEMGROUPCODE': '',
            'SALESUNITSYMBOL': '',
            'SEARCHNAME': '',
            'SERVICETYPE': 'NotSpecified',
            'STORAGEDIMENSIONGROUPNAME': 'WMS',
            'TRACKINGDIMENSIONGROUPNAME': 'BATCH',
            'UPPERWARRANTABLEPRICERANGELIMIT': '',
            'VARIANTCONFIGURATIONTECHNOLOGY': '',
            'WARRANTABLEPRICERANGEBASETYPE': '',
            'WARRANTYDURATIONTIME': '',
            'WARRANTYDURATIONTIMEUNIT': '',
        })

        output_file = 'RELEASED_PRODUCT_CREATE.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


# ---------------------------------------------------------------------------
# DM ITEM BAR CODE
# ---------------------------------------------------------------------------

BARCODE_MAPPING = {
    'B9NNN2': 'EAN-13',
    'B9NHNB': 'CIP-7',
    'B9OIN2': 'GTIN',
    'B9UKC1': 'EXTCODE',
}


def process_item_bar_code(products_file):
    try:
        products_data = pd.read_excel(products_file, dtype=str)
        products_data.columns = products_data.columns.str.upper()

        long_df = products_data.reset_index().melt(
            id_vars=['index', 'B9N6N2'],
            value_vars=list(BARCODE_MAPPING.keys()),
            var_name='SOURCE_COL',
            value_name='BARCODE',
        )
        long_df['BARCODE'] = long_df['BARCODE'].fillna('').str.strip()
        long_df = long_df[~long_df['BARCODE'].isin(['', '0'])]
        long_df = long_df.sort_values('index', kind='stable').reset_index(drop=True)

        output_data = pd.DataFrame({
            'ITEMNUMBER': '',
            'PRODUCTQUANTITYUNITSYMBOL': 'PCS',
            'BARCODESETUPID': long_df['SOURCE_COL'].map(BARCODE_MAPPING),
            'ISDEFAULTDISPLAYEDBARCODE': 'NO',
            'ISDEFAULTPRINTEDBARCODE': 'NO',
            'ISDEFAULTSCANNEDBARCODE': 'NO',
            'PRODUCTQUANTITY': 0,
            'BARCODE': long_df['BARCODE'],
            'PRODUCTNUMBER': long_df['B9N6N2'].fillna('').str.strip(),
        })

        output_file = 'item_code_bar.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {e}") from e


# ---------------------------------------------------------------------------
# DM APPROVED VENDOR LIST / DEFAULT ORDER SETTINGS
# ---------------------------------------------------------------------------

def process_approved_vendor_list_by_products(products_file):
    try:
        products_data = pd.read_excel(products_file)
        products_data.columns = products_data.columns.str.upper()

        output_data = pd.DataFrame({
            'PRODUCTNUMBER': products_data['B9N6N2'],
            'ITEMNUMBER': '',
            'APPROVEDVENDORACCOUNTNUMBER': products_data['NUMSUPP'],
            'EFFECTIVEDATE': '',
            'EXPIRATIONDATE': '',
        })
        output_file = 'PRODUIT_approved_vendor_list_by_products.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {str(e)}")


def process_default_order_settings(products_file, transco_file):
    try:
        products_data = pd.read_excel(products_file, dtype=str)
        transco_pays = pd.read_excel(transco_file, dtype=str, sheet_name="PaysOrigine")
        products_data.columns = products_data.columns.str.strip().str.upper()
        transco_pays.columns = transco_pays.columns.str.upper()

        output_data = pd.DataFrame({
            'ITEMNUMBER': '',
            'ISPROCUREMENTPROCESSINGSTOPPED': '',
            'ISPROCUREMENTSITEMANDATORY': 'NO',
            'ISPROCUREMENTUSINGWORKINGDAYS': 'NO',
            'ISPROCUREMENTWAREHOUSEMANDATORY': 'NO',
            'MAXIMUMPROCUREMENTORDERQUANTITY': '',
            'MINIMUMPROCUREMENTORDERQUANTITY': '',
            'PROCUREMENTLEADTIMEDAYS': '',
            'PROCUREMENTQUANTITYMULTIPLES': '',
            'PROCUREMENTSITEID': '',
            'PROCUREMENTWAREHOUSEID': '',
            'ISSALESPROCESSINGSTOPPED': '',
            'STANDARDSALESORDERQUANTITY': '',
            'MULTIPLESALESORDERQUANTITY': '',
            'MINIMUMSALESORDERQUANTITY': '',
            'MAXIMUMSALESORDERQUANTITY': '',
            'PRODUCTNUMBERLAG': products_data['B9N6N2'],
        })

        output_file = 'PRODUIT_default_order_settings.xlsx'
        output_data.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        raise Exception(f"Erreur lors de la transformation: {e}") from e


# ---------------------------------------------------------------------------
# Gradio interface
# ---------------------------------------------------------------------------

def _single_file_tab(title, fn, input_label, button_label, output_label):
    with gr.Tab(title):
        inp = gr.File(label=input_label)
        btn = gr.Button(button_label)
        out = gr.File(label=output_label)
        btn.click(fn=fn, inputs=inp, outputs=out)


def _two_file_tab(title, fn, input_label, transco_label, button_label, output_label):
    with gr.Tab(title):
        inp = gr.File(label=input_label)
        transco = gr.File(label=transco_label)
        btn = gr.Button(button_label)
        out = gr.File(label=output_label)
        btn.click(fn=fn, inputs=[inp, transco], outputs=out)


with gr.Blocks(title="Articles Transcoding") as iface:
    gr.Markdown("Articles Transcoding file tools.")

    _single_file_tab("Products", process_products_file,
                     "Upload Products File", "Transcode Products File",
                     "Processed Products File")
    _two_file_tab("Released Products", process_released_products,
                  "Upload Products File", "Upload Transco file for released products",
                  "Transcode released Products File", "Processed released Products File")
    _single_file_tab("Product bar code", process_item_bar_code,
                     "Upload Products File", "Transcode item bar code Products File",
                     "Processed item bar code File")
    _single_file_tab("Products Translation", process_products_translation,
                     "Upload Products Translation File", "Transcode Products Translation File",
                     "Processed Products Translation File")
    _single_file_tab("Released Products Create", transform_released_products_create,
                     "Upload Products released create File", "Transcode Products released create File",
                     "Processed Products released create File")
    _single_file_tab("Products Category Assignment", process_products_category_assignment,
                     "Upload Products Category Assignment File",
                     "Transcode Products Category Assignment File",
                     "Processed Products Category Assignment File")
    _single_file_tab("Approved vendor list by products", process_approved_vendor_list_by_products,
                     "Upload Approved vendor list by products File",
                     "Transcode Approved vendor list by products File",
                     "Processed Approved vendor list by products File")
    _two_file_tab("Default order settings", process_default_order_settings,
                  "Upload default order settings Products File",
                  "Upload Transco file for default order settings products",
                  "Transcode default order settings Products File",
                  "Processed default order settings Products File")
    _single_file_tab("External item descriptions for customers",
                     process_products_external_items_for_customers,
                     "Upload External item descriptions for customers File",
                     "Transcode External item descriptions for customers File",
                     "Processed External item descriptions for customers File")


if __name__ == "__main__":
    iface.queue().launch(css=".footer {visibility: hidden;}")
