import streamlit as st
import openpyxl
from deep_translator import GoogleTranslator
import io
import pandas as pd

# --- PODEŠAVANJE STRANICE ---
st.set_page_config(page_title="Excel Prevodilac (SR -> BN)", layout="centered")

st.title("🇧🇩 Excel Prevodilac: Srpski -> Bengalski")
st.markdown("""
Ova aplikacija prevodi Excel sheet zadržavajući formatiranje, boje i mergovana polja.
**Uputstvo:**
1. Uploaduj .xlsx fajl
2. Izaberi sheet
3. Klikni na Start
""")

# --- FUNKCIJA ZA PREVOD ---
def translate_excel(file, sheet_name):
    # Učitavanje u memoriju
    wb = openpyxl.load_workbook(file)
    
    # Kreiranje kopije sheeta
    if f"{sheet_name}_Bengali" in wb.sheetnames:
        # Brišemo stari ako postoji da ne pravi duplikate
        del wb[f"{sheet_name}_Bengali"]
        
    source = wb[sheet_name]
    target = wb.copy_worksheet(source)
    target.title = f"{sheet_name[:20]}_Bengali" # Skraćujemo ime zbog limita
    
    translator = GoogleTranslator(source='sr', target='bn')
    
    # Sakupljanje ćelija za prevod
    cells_to_translate = []
    
    # Iteracija kroz redove
    # Koristimo progress bar placeholder
    progress_text = "Skeniram fajl..."
    my_bar = st.progress(0, text=progress_text)
    
    total_cells = 0
    for row in target.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                # Preskačemo čiste brojeve koji su formatirani kao tekst
                if not cell.value.strip().isdigit():
                    cells_to_translate.append(cell)
    
    total_items = len(cells_to_translate)
    st.info(f"Pronađeno {total_items} polja sa tekstom. Počinjem prevod...")
    
    # Cache za prevode da ne trošimo vreme na iste reči
    translation_cache = {}
    
    # Glavna petlja prevoda
    for i, cell in enumerate(cells_to_translate):
        text = cell.value.strip()
        
        # Ažuriranje progress bara na svakih 5%
        if i % 10 == 0:
            percent = int((i / total_items) * 100)
            my_bar.progress(percent, text=f"Prevodim: {text[:20]}...")
            
        if text in translation_cache:
            cell.value = translation_cache[text]
        else:
            try:
                translated = translator.translate(text)
                translation_cache[text] = translated
                cell.value = translated
            except Exception as e:
                continue # Ako pukne jedna reč, nastavi dalje

    my_bar.progress(100, text="Završeno!")
    
    # Čuvanje u memorijski buffer (ne na disk)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

# --- INTERFEJS ---
uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Samo učitamo imena sheetova (brzo je)
        wb_temp = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
        sheet_names = wb_temp.sheetnames
        wb_temp.close()
        
        selected_sheet = st.selectbox("Koji sheet želiš da prevedeš?", sheet_names)
        
        if st.button("🚀 Pokreni Prevod"):
            with st.spinner('Radim... Ovo može potrajati par minuta zavisno od veličine fajla.'):
                # Pozivamo funkciju
                processed_data = translate_excel(uploaded_file, selected_sheet)
                
                st.success("Prevod je gotov!")
                
                # Dugme za download
                st.download_button(
                    label="📥 Preuzmi prevedeni fajl",
                    data=processed_data,
                    file_name=f"PREVEDENO_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Došlo je do greške pri učitavanju fajla: {e}")
