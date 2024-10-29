import os
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.core.files.storage import FileSystemStorage
import openpyxl  # For handling Excel
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ParseError

class FileUploadView(APIView):
    def post(self, request):
        try:
            if 'file' not in request.FILES:
                return Response({'error': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)

            xml_file = request.FILES['file']
            
            # Validate if the uploaded file is an XML file
            if not xml_file.name.endswith('.xml'):
                return Response({'error': 'Invalid file format. Only XML files are allowed.'}, status=status.HTTP_400_BAD_REQUEST)

            # Save the uploaded file
            fs = FileSystemStorage()
            filename = fs.save(xml_file.name, xml_file)
            file_path = fs.path(filename)

            # Process the XML file
            xlsx_path = self.process_tally_xml(file_path)

            # Return success message
            return Response({'message': f'File {os.path.basename(xlsx_path)} created successfully'}, status=status.HTTP_201_CREATED)
        
        except ParseError:
            return Response({'error': 'Invalid XML structure or corrupted file.'}, status=status.HTTP_400_BAD_REQUEST)
        except ET.ElementTree as e:
            return Response({'error': f'Error parsing XML: {str(e)}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
        except Exception as e:
            return Response({'error': f'An unexpected error occurred: {str(e)}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    def process_tally_xml(self, xml_file):
        """
        this function convert the .xml file data into .xlsx file
        input: xml_file
        """
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Create a new Excel workbook
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = 'Tally Receipts'

            sheet.append(['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount', 'Amount', 'Particulars', 'Vch Type', 'Amount Verified'])

            for voucher in root.findall('.//VOUCHER[@VCHTYPE="Receipt"]'):
                date = voucher.find('DATE').text
                voucher_type = voucher.attrib['VCHTYPE']
                party_ledger_name = voucher.find('PARTYLEDGERNAME').text
                voucher_number = voucher.find('VOUCHERNUMBER').text

                for ledger in voucher.findall('.//ALLLEDGERENTRIES.LIST'):
                    ledger_name = ledger.find('LEDGERNAME').text
                    amount = ledger.find('AMOUNT').text

                    if ledger_name == party_ledger_name:
                        sheet.append([date, 'Parent', voucher_number, 'NA', 'NA', 'NA', ledger_name, 'NA', amount, ledger_name, voucher_type, 'YES'])
                    else:
                        sheet.append([date, 'other', voucher_number, 'NA', 'NA', 'NA', ledger_name, 'NA', amount, ledger_name, voucher_type, 'NA'])

                    if ledger.findall('.//BILLALLOCATIONS.LIST'):
                        for bill_allocation in ledger.findall('.//BILLALLOCATIONS.LIST'):
                            if bill_allocation.find("AMOUNT") is not None:
                                bill_amount = bill_allocation.find("AMOUNT").text
                                ref_no = bill_allocation.find('NAME').text
                                ref_type = bill_allocation.find('BILLTYPE').text

                                sheet.append([date, 'child', voucher_number, ref_no, ref_type, '', ledger_name, bill_amount, 'NA', ledger_name, voucher_type, 'NA'])

            output_filename = os.path.splitext(xml_file)[0] + '_processed.xlsx'
            workbook.save(output_filename)

            return output_filename
        
        except ET.ParseError:
            raise ParseError("Error parsing XML file")
        except Exception as e:
            raise Exception(f"An error occurred while processing the file: {str(e)}")
