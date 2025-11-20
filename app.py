import streamlit as st
import openpyxl
from deep_translator import GoogleTranslator
import io

# --- KONFIGURACIJA ---
st.set_page_config(page_title="Single Sheet Translator", layout="centered")

st.title("🇧🇩 Excel Prevodilac (Samo 1 Sheet)")
st.markdown("""
Ova aplikacija:
1. Uzima jedan sheet koji izabereš.
2. Pretvara sve formule u tekst (da se ne pokvare).
3. Prevodi tekst na bengalski.
4. **Briše sve ostale sheetove** i daje ti čist fajl.
""")

def translate_single_sheet(file, sheet_name):
    # 1. Učitavanje fajla (data_only=True SKIDA FORMULE i ostavlja vrednosti)
    # Ovo je ključno da prevod ne bi brljao formule
    wb = openpyxl.load_workbook(file, data_only=True)
    
    if sheet_name not in wb.sheetnames:
        return None
        
    # Radimo direktno na originalnom sheetu jer ćemo ostale obrisati
    ws = wb[sheet_name]
    ws.title = "Bengali_Prevod" # Menjamo ime sheeta
    
    # 2. BRISANJE OSTALIH SHEETOVA (Izolacija)
    # Prolazimo kroz sva imena i brišemo sve što nije naš sheet
    for name in wb.sheetnames:
        if name != "Bengali_Prevod":
            del wb[name]
            
    # Sada je u 'wb' ostao samo jedan sheet. Njega prevodimo.
    
    # 3. Optimizovano skeniranje teksta
    unique_texts = set()
    cells_map = [] 

    progress_bar = st.progress(0, text="Analiziram sadržaj...")
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                # Uslov: Nije prazno i nije samo broj
                if val and not val.isdigit():
                    unique_texts.add(val)
                    cells_map.append(cell)

    total_unique = len(unique_texts)
    st.caption(f"Pronađeno {total_unique} fraza za prevod.")
    
    if total_unique > 0:
        # 4. Prevod
        translator = GoogleTranslator(source='sr', target='bn')
        translation_dict = {}
        
        progress_bar.progress(10, text="Prevodim...")
        unique_list = list(unique_texts)
        
        for i, text in enumerate(unique_list):
            # Update bara ređe (svakih 5%) da ne koči
            if i % 5 == 0:
                percent = 10 + int((i / total_unique) * 80)
                progress_bar.progress(percent, text=f"Prevodim: {text[:15]}...")
            
            try:
                translated = translator.translate(text)
                translation_dict[text] = translated
            except:
                translation_dict[text] = text
                
        # 5. Primena prevoda
        progress_bar.progress(95, text="Upisujem podatke...")
        for cell in cells_map:
            original = cell.value.strip()
            if original in translation_dict:
                cell.value = translation_dict[original]

    # 6. Čuvanje
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    progress_bar.progress(100, text="Gotovo!")
    
    return output

# --- INTERFEJS ---
uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Samo čitanje strukture
        wb_temp = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
        sheet_names = wb_temp.sheetnames
        wb_temp.close()
        
        selected_sheet = st.selectbox("Koji sheet želiš da izdvojiš i prevedeš?", sheet_names)
        
        if st.button("🚀 Prevedi i Izdvoj"):
            with st.spinner('Obrađujem...'):
                result = translate_single_sheet(uploaded_file, selected_sheet)
                
                if result:
                    st.success("Završeno!")
                    st.download_button(
                        label="📥 Preuzmi XLSX (Samo 1 sheet)",
                        data=result,
                        file_name=f"{selected_sheet}_Bengali.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Došlo je do greške.")
                    
    except Exception as e:
        st.error(f"Greška: {e}")
