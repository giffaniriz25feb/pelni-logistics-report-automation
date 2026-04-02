from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
import pandas as pd
import os
import time
import datetime

# === CONFIGURABLE PATH ===
edge_driver_path = r"C:\Users\user\lib\edgedriver_win64\msedgedriver.exe"
photo_dir = r"C:\Users\user\Lognus 5 Automation\Out_Lognus1_5"

# === Helper Functions ===

def highlight(driver, element):
    driver.execute_script("arguments[0].style.border='3px solid red'", element)

def wait_for_overlay_to_disappear(driver, timeout=10):
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "blockUI"))
        )
        print("✅ Overlay hilang, lanjutkan.")
    except:
        print("⚠️ Timeout: overlay masih muncul.")

def select2_dropdown(driver, dropdown_container_id, search_text, description, allow_fail=False):
    for attempt in range(3):
        try:
            print(f"🔍 Attempting dropdown: {description}")
            wait_for_overlay_to_disappear(driver)

            container = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, dropdown_container_id))
            )
            highlight(driver, container)
            container.click()
            time.sleep(1)

            input_box = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field"))
            )
            input_box.clear()
            input_box.send_keys(search_text)
            time.sleep(1.2)

            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "select2-results__option"))
            )
            time.sleep(0.3)
            options = driver.find_elements(By.CLASS_NAME, "select2-results__option")
            if options:
                highlight(driver, options[0])
                time.sleep(0.4)
                options[0].click()
                print(f"✅ Selected '{search_text}' for {description}")
                return True
            else:
                print(f"❌ No options found for {description}")
                return False

        except Exception as e:
            print(f"❌ Failed dropdown {description}: {e}")
            time.sleep(1)

    if allow_fail:
        return False
    raise Exception(f"Dropdown failed after 3 attempts: {description}")

def upload_file(driver, input_id, file_path, description):
    try:
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, input_id))
        )
        highlight(driver, file_input)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"❌ File not found: {file_path}")
        driver.execute_script("arguments[0].style.display = 'block';", file_input)
        file_input.send_keys(file_path)
        print(f"📎 Uploaded {description}: {os.path.basename(file_path)}")
    except Exception as e:
        print(f"❌ File upload failed for {description}: {e}")

def check_popup(driver):
    try:
        title = driver.find_element(By.CLASS_NAME, "swal2-title").text.lower()
        message = driver.find_element(By.CLASS_NAME, "swal2-html-container").text.lower()
        full_text = f"{title} {message}"

        if "gagal" in full_text:
            print(f"❌ SweetAlert Message: {full_text}")
            return "Gagal"
        elif "success" in full_text:
            print(f"✅ SweetAlert Message: {full_text}")
            return "Success"
        else:
            print(f"ℹ️ SweetAlert Message: {full_text}")
            return "Unknown"
    except Exception as e:
        print(f"⚠️ Gagal membaca popup: {e}")
        return "Unknown"

# === Main Script ===

def run_script(xlsx_path):
    df = pd.read_excel(xlsx_path)

    options = Options()
    options.use_chromium = True
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)

    driver.get("https://cis.pelni.co.id/transaction/formTransaction")
    print("⏳ Please login manually... then press Enter.")
    input("⏸️  Press Enter to continue...")

    success_inputs = []
    input_failures = []
    dropdown_failures = []

    consecutive_failures = 0
    total_rows = len(df)
    start_time = time.time()

    for index, row in df.iterrows():
        print(f"\n➡️ Processing index {index + 1}/{total_rows}...")
        elapsed = time.time() - start_time
        est_total = (elapsed / (index + 1)) * total_rows
        est_remaining = est_total - elapsed
        progress = (index + 1) / total_rows * 100
        print(f"⏳ Progress: {progress:.2f}% | ETA: {datetime.timedelta(seconds=int(est_remaining))}")

        if str(row.get("Automate", "")).strip().lower() == "no":
            print("⏩ Skipping: Not marked for automation.")
            continue

        try:
            kontainer_ok = select2_dropdown(driver, "select2-container-container", str(row['Kontainer']), "#container", allow_fail=True)
            if not kontainer_ok:
                dropdown_failures.append((row['Kontainer'], 'Kontainer'))
                continue

            aktivitas_id = "aktivitas-in" if row['Aktivitas'].lower() == "in" else "aktivitas-out"
            aktivitas_box = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, aktivitas_id)))
            highlight(driver, aktivitas_box)
            wait_for_overlay_to_disappear(driver)
            aktivitas_box.click()

            tujuan_ok = select2_dropdown(driver, "select2-tujuan-container", row['Dari / Ke'], "#tujuan", allow_fail=True)
            if not tujuan_ok:
                dropdown_failures.append((row['Kontainer'], 'Tujuan'))
                continue

            kapal_ok = select2_dropdown(driver, "select2-kapal-container", row['Kapal'], "#kapal", allow_fail=True)
            if not kapal_ok:
                dropdown_failures.append((row['Kontainer'], 'Kapal'))
                continue

            voyage_select = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "voyage")))
            highlight(driver, voyage_select)
            wait_for_overlay_to_disappear(driver)
            Select(voyage_select).select_by_visible_text(str(row['Voyage']))

            muatan = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "muatan")))
            highlight(driver, muatan)
            Select(muatan).select_by_visible_text("Full" if row['Muatan'].lower() != "empty" else "Empty")

            try:
                jenis = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "grade")))
                highlight(driver, jenis)
                Select(jenis).select_by_visible_text(row['Jenis Kontainer'])
            except:
                pass

            photo_path = os.path.join(photo_dir, row['Foto Keseluruhan'])
            upload_file(driver, "foto_keseluruhan", photo_path, "Photo")

            pdf_path = os.path.join(os.path.dirname(xlsx_path), row['Dokumen Pendukung'])
            upload_file(driver, "dok", pdf_path, "PDF")

            simpan = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-simpan")))
            highlight(driver, simpan)
            wait_for_overlay_to_disappear(driver)
            simpan.click()

            for _ in range(30):
                try:
                    # Pastikan popup muncul
                    WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "swal2-title"))
                    )

                    # Baca judul dan isi popup
                    title = driver.find_element(By.CLASS_NAME, "swal2-title").text.lower()
                    message = driver.find_element(By.CLASS_NAME, "swal2-html-container").text.lower()
                    full_text = f"{title} {message}"
                    print(f"📢 SweetAlert Message: {full_text}")

                    # Tentukan status berdasarkan isi pesan
                    if "gagal" in full_text:
                        popup_status = "Gagal"
                    elif "success" in full_text:
                        popup_status = "Success"
                    else:
                        popup_status = "Unknown"

                    # Klik tombol OK setelah baca
                    popup_button = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.CLASS_NAME, "swal2-confirm"))
                    )
                    popup_button.click()
                    break
                except:
                    time.sleep(0.5)
            else:
                print("⚠️ SweetAlert2 popup not found")
                popup_status = "Unknown"


            if popup_status == "Gagal":
                input_failures.append(row['Kontainer'])
                consecutive_failures += 1
            else:
                success_inputs.append(row['Kontainer'])
                consecutive_failures = 0

            if consecutive_failures >= 6:
                input_failures.append(f"{row['Kontainer']} (consec skip)")
                consecutive_failures = 0

            time.sleep(1)

        except KeyboardInterrupt:
            print("🛑 Interrupted by user.")
            break

        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            input_failures.append(row['Kontainer'])
            consecutive_failures += 1

    # Simpan hasil ke Excel
    hasil_df = pd.DataFrame({
        "Kontainer": (
            success_inputs +
            [x for x in input_failures if x not in success_inputs] +
            [x[0] for x in dropdown_failures]
        ),
        "Status": (
            ["Sukses"] * len(success_inputs) +
            ["Gagal Input"] * len([x for x in input_failures if x not in success_inputs]) +
            [f"Gagal Dropdown - {x[1]}" for x in dropdown_failures]
        )
    })

    output_file = os.path.join(os.path.dirname(xlsx_path), "hasil_log_input.xlsx")
    hasil_df.to_excel(output_file, index=False)
    print(f"\n📁 Hasil log disimpan di: {output_file}")
    print("🏁 Script selesai.")

# Run the script
run_script(r"C:\Users\user\OneDrive\Documents\Rangga\Kuliah S1 Statistika\KP\Lognus 5 Automation\Manifest Logistik Nusantara 1 Voyage 5 Out.xlsx")
