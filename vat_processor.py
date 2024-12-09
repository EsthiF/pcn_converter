import pandas as pd
from datetime import datetime
import re
import numpy as np

class VATReportProcessor:
    def __init__(self):
        self.transaction_format = {
            'Entry Type': ('A', 1),
            'VAT identification number': ('N', 9),
            'Invoice Date': ('N', 8),
            'Reference group': ('A', 4),
            'Reference number': ('N', 9),
            'Total VAT in invoice': ('N', 9),
            'Invoice total not incl. VAT': ('N', 10),
            'Space for future data': ('N', 9)
        }

    def format_field(self, value, format_type, length, add_sign=False):
        if pd.isna(value) or value == '' or value is None:
            if format_type == 'N' or format_type == 'A':
                return ('0' * length) if not add_sign else ('+' + '0' * length)

        if format_type == 'N':
            try:
                # Handle numeric values
                if isinstance(value, (int, float)):
                    num_value = int(round(float(value)))  
                else:
                    # Try to convert string to number
                    num_value = int(float(str(value).replace(',', '')))
                
                is_negative = num_value < 0
                value_str = str(abs(num_value))
            except:
                # If conversion fails, treat as string
                value_str = ''.join(filter(str.isdigit, str(value)))
                is_negative = False

            if not value_str:
                return ('0' * length) if not add_sign else ('+' + '0' * length)
            
            # Take last 'length' characters or pad with zeros
            if len(value_str) > length:
                value_str = value_str[-length:]
            else:
                value_str = value_str.zfill(length)
            
            if add_sign:
                return f"{'-' if is_negative else '+'}{value_str}"
            return value_str
            
        elif format_type == 'A':
            value_str = str(value)
            if value_str.lower() == 'nan':
                return '0' * length
            return value_str[:length].ljust(length, '0')

    def process_file(self, input_excel):
        try:
            # Read company info first
            company_info = pd.read_excel(input_excel, nrows=4)
            vat_number = str(company_info.iloc[1, 1]).strip().split('.')[0]
            year = str(company_info.iloc[2, 1]).strip().split('.')[0]
            month = str(company_info.iloc[3, 1]).strip().split('.')[0].zfill(2)
            report_month = f"{year}{month}"
            
            # Read the transaction data, skipping the headers
            df = pd.read_excel(input_excel, skiprows=8)
            
            output_filename = 'PCN874.TXT'
            
            with open(output_filename, 'w') as f:
                # Write header record (O record)
                header_line = 'O'
                header_line += self.format_field(vat_number, 'N', 9)
                header_line += self.format_field(report_month, 'N', 6)
                header_line += self.format_field('1', 'N', 1)
                header_line += self.format_field(datetime.now().strftime('%Y%m%d'), 'N', 8)
                
                # Calculate totals (skipping any non-numeric rows)
                df_numeric = df[df.iloc[:, 5].apply(lambda x: str(x).replace('.', '').isdigit())]
                total_vat = int(df_numeric.iloc[:, 5].astype(float).sum())
                total_amount = int(df_numeric.iloc[:, 6].astype(float).sum())
                
                header_line += self.format_field(total_amount, 'N', 11, True)
                header_line += self.format_field(total_vat, 'N', 9, True)
                f.write(header_line + '\n')
                
                # Write transaction records
                for _, row in df.iterrows():
                    if pd.notna(row.iloc[0]) and row.iloc[0] != 'Entry Type':
                        trans_line = ''  # Start with no prefix, use Entry Type directly
                        trans_line += self.format_field(row.iloc[0], 'A', 1)
                        trans_line += self.format_field(row.iloc[1], 'N', 9)
                        trans_line += self.format_field(row.iloc[2], 'N', 8)
                        trans_line += self.format_field(row.iloc[3], 'A', 4)
                        trans_line += self.format_field(row.iloc[4], 'N', 9)
                        trans_line += self.format_field(row.iloc[5], 'N', 9, True)
                        trans_line += self.format_field(row.iloc[6], 'N', 10, True)
                        trans_line += self.format_field(row.iloc[7] if len(row) > 7 else None, 'N', 9)
                        f.write(trans_line + '\n')
                
                # Write closing record (H record)
                closing_line = 'H'
                closing_line += self.format_field(vat_number, 'N', 9)
                f.write(closing_line + '\n')
                
            print(f"\nSuccess! Generated {output_filename}")
            
        except Exception as e:
            print(f"\nError processing file: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            traceback.print_exc()
            raise

# Example usage
if __name__ == "__main__":
    processor = VATReportProcessor()
    processor.process_file("vat_report.xlsx")