import os
import shutil
import time
from glob import glob
from openai import OpenAI
from oletools.olevba import VBA_Parser

# OpenAI API nyckel
api_nyckel = ""
nyckel = OpenAI(api_key=api_nyckel)  

# Mapp där .docm filerna är sparade
mapp = ""

# Mapp dit analyserade filer flyttas
destinations_mapp = ""

# Fil för svaren
output_fil_path = os.path.join(mapp, "alla_svar.txt")

while True:
    # Hitta alla .doc filer
    docm_fil = os.path.join(mapp, '*.docm')
    alla_doc_filer = glob(docm_fil)
    
    # Hantera makros
    for varje_fil in alla_doc_filer:
        vba_parser = VBA_Parser(varje_fil)

        alla_makro = ""
        # Se om det finns makros i filen
        if vba_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                alla_makro += f"Macro in: {filename}\n{vba_code}\n{'-' * 40}\n"

        # Flytta analyserad fil 
        shutil.move(varje_fil, os.path.join(destinations_mapp, os.path.basename(varje_fil)))
        
        # Skicka makros till GPT
        if alla_makro:
            svar = nyckel.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "user", "content": f"Act as an IT security analyst, analyze the following VBA code to determine its potential for harmful actions. Provide a percentage estimate of how likely it is to be malicious:\n\n{alla_makro}"}
                ]
            )

            # Skriv ut GPTs svar till .txt filen
            with open(output_fil_path, 'a') as output_fil:
                base_name = os.path.basename(varje_fil)
                output_fil.write(f"File: {base_name}\n")
                output_fil.write("GPT's analysis:\n")
                output_fil.write(svar.choices[0].message.content + "\n")
                output_fil.write("=" * 80 + "\n\n")  


        vba_parser.close()

        # Vänta en minut efter varje fil för att säkerhetställa att man inte kommer upp i token limit
        time.sleep(60)
        break
