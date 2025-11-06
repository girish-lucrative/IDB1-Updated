from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (StaleElementReferenceException, 
                                      TimeoutException,
                                      NoSuchElementException)
from datetime import datetime
from datetime import timedelta
import time
import os
import logging
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import winsound

class CertificateBot:
    def __init__(self, username, password, excel_data, download_folder):
        self.username = username
        self.password = password
        self.excel_data = excel_data
        # self.required_files_folder = required_files_folder
        self.download_folder = download_folder
        self.driver = None
        self.current_index = 0
        
        
        # self.required_docs = os.path.abspath(required_files_folder)

    def process_all_certificates(self):
        try:
            if not self.excel_data:
                return {"success": False, "message": "No data found in Excel file"}
            
            # Process all certificates
            while self.current_index < len(self.excel_data):
                
                full_login = self.current_index == 0
                result = self._process_certificate(full_login=full_login)
                if not result.get('success'):
                    return result
            
                        
        except Exception as e:
            return {"success": False, "message": f"Error processing certificates: {e}"}
        finally:
            self.close_browser()

    def _process_certificate(self, full_login=False):
        try:
            row = self.excel_data[self.current_index]

            
            if full_login:
                self.start_browser()
                login_result = self.login()
              
            
            fill_result = self.fill_certificate()
            if not fill_result.get('success'):
                return fill_result
            
            self.current_index += 1
            return {"success": True, "message": f"Processed  {row}"}
            
        except Exception as e:
            return {"success": False, "message": f"Error processing certificate: {e}"}

    def start_browser(self):
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("detach", True)
        # Do NOT add headless option, keep browser visible
        # options.add_argument("--headless")  # (do not use)

        # profile_dir = os.path.join(self.download_folder, "chrome_profile")
        # if not os.path.exists(profile_dir):
        #     os.makedirs(profile_dir)
        # options.add_argument(f"--user-data-dir={profile_dir}")
        
        
        options.add_argument("--disable-gpu")
        self.driver = webdriver.Chrome(options=options)
        self.driver.maximize_window()

    def login(self):
        try:
            self.driver.get("https://www.icegate.gov.in")
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Login / Sign Up"))).click()
            print("click to login")

            # Wait briefly to ensure the new tab is opened
            time.sleep(4)
            self.driver.refresh()
            # Get all window handles
            tabs = self.driver.window_handles
            
            # Switch to the newest tab (usually the last in the list)
            self.driver.switch_to.window(tabs[-1])
            
            # Now you're in the new tab; continue your automation
            print("Switched to new tab:", self.driver.current_url)

            # self.driver.execute_script("window.scrollBy(0, 50);")
            # self.driver.refresh()

            WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located((By.ID, "icegateId")))
            print("wait for icegateID")
            
            self.driver.find_element(By.ID, "icegateId").send_keys(self.username)
            self.driver.find_element(By.ID, "password").send_keys(self.password)
            original_url = self.driver.current_url
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LOGIN']"))).click()
            

            # Wait until the URL changes (max 5 minutes)
            WebDriverWait(self.driver, 100).until(EC.url_changes(original_url))

            print("change URL")
            
            
            return {"success": True, "message": "Login successful"}
            
        except Exception as e:
            return {"success": False, "message": f"Login failed: {e}"}
            
    def fill_certificate(self):
        try:
            print(1)
            time.sleep(1)
            row = self.excel_data[self.current_index]
            if self.current_index==0:

                # services = self.driver.find_element(By.XPATH, "//h5[normalize-space()='Services']")
                # print("services")
                # self.driver.execute_script("arguments[0].click();", services)
                # print("click on services")


                try:
                   services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
                   
                   print("services")
                   self.driver.execute_script("arguments[0].click();", services)
                   print("click on services")
                except:
                    print("not work 5")
                    services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                    print("services")
                    self.driver.execute_script("arguments[0].click();", services)
                    print("click on services")
                
    
                self.driver.execute_script("window.scrollBy(0, 100);")
                time.sleep(1)
    
                
    
                wait = WebDriverWait(self.driver, 100)
    
                # Click the chevron corresponding to the "E-Payment" node
                # Enquiries = wait.until(EC.element_to_be_clickable((
                #     By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button/span/mat-icon'
                # )))
                Enquiries = wait.until(EC.element_to_be_clickable((
                    By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
                )))
                Enquiries.click()
                print("click on Enquiries")

                icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                    By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
                )))
                icegate_enquiry_service.click()
                print("click on Icegate Enquiry services")

                drawback_disbursement_status = wait.until(EC.element_to_be_clickable((
                    By.XPATH, '/html/body/app-root/app-layout/div/div[2]/app-enquiries-layout/app-available-enquiries/div/div/div[1]/div/div[7]/div/div/span'
                )))
                drawback_disbursement_status.click()
                print("click on drawback disbursement status")

                global_start_date = None
                global_end_date = None
                
                for idx, row in enumerate(self.excel_data):
                    location_code = str(row.get("Realisation Port", "")).strip()
                    print(f"\n📍 Processing Port {idx+1}: {location_code}")
                
                    def parse_date_safe(date_value):
                        if pd.isna(date_value):
                            return None
                        if isinstance(date_value, str):
                            try:
                                return datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S")
                            except ValueError:
                                try:
                                    return datetime.strptime(date_value, "%Y-%m-%d")
                                except:
                                    return None
                        return date_value if isinstance(date_value, datetime) else None
                
                    # --- For first port, read dates from Excel ---
                    if idx == 0:
                        start_date_obj = parse_date_safe(row.get("Realisation Start Date"))
                        end_date_obj = parse_date_safe(row.get("Realisation End Date"))
                
                        if not start_date_obj or not end_date_obj:
                            print(f"⚠️ Invalid or missing date(s) for first port: {location_code}")
                            break
                
                        # Save for reuse
                        global_start_date = start_date_obj
                        global_end_date = end_date_obj
                        print(f"🗓️ Base Date Range Set: {global_start_date.strftime('%d-%m-%Y')} → {global_end_date.strftime('%d-%m-%Y')}")
                
                    # --- For later ports, reuse first port’s dates ---
                    else:
                        start_date_obj = global_start_date
                        end_date_obj = global_end_date
                
                    current_from = start_date_obj
                
                    # === Loop through all 30-day windows ===
                    while current_from <= end_date_obj:
                        current_to = current_from + timedelta(days=29)
                        if current_to > end_date_obj:
                            current_to = end_date_obj
                
                        from_str = current_from.strftime("%d-%m-%Y")
                        to_str = current_to.strftime("%d-%m-%Y")
                
                        # --- Fill 'Date From' ---
                        self.driver.execute_script("""
                            var el = document.querySelector('input[formcontrolname="startDate"]');
                            if (el) {
                                el.removeAttribute('readonly');
                                el.value = arguments[0];
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                            }
                        """, from_str)
                
                        # --- Fill 'To' date ---
                        self.driver.execute_script("""
                            var el = document.querySelector('input[formcontrolname="endDate"]');
                            if (el) {
                                el.removeAttribute('readonly');
                                el.value = arguments[0];
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                            }
                        """, to_str)
                
                        print(f"🗓️ Date Range: {from_str} → {to_str}")
                
                        # --- Select Port ---
                        # try:
                        #     time.sleep(1)
                        #     port_dropdown = wait.until(
                        #         EC.element_to_be_clickable((By.XPATH, "//ng-select[contains(@formcontrolname,'location') or contains(.,'Select Port')]"))
                        #     )
                        #     port_dropdown.click()
                        #     time.sleep(1)
                
                        #     search_input = wait.until(
                        #         EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                        #     )
                        #     search_input.clear()
                        #     search_input.send_keys(location_code)
                        #     time.sleep(2)
                
                        #     first_option = wait.until(
                        #         EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                        #     )
                        #     self.driver.execute_script("arguments[0].click();", first_option)
                        #     time.sleep(1)
                        #     print(f"✅ Selected location: {location_code}")
                        # except Exception as e:
                        #     print(f"⚠️ Port selection failed for {location_code}: {e}")
                        #     break
                        try:
                            # Wait for dropdown to be clickable in DOM
                            port_dropdown = wait.until(
                                EC.presence_of_element_located(
                                    (By.XPATH, "//ng-select[contains(@formcontrolname,'locationCode') or contains(.,'Select Port')]")
                                )
                            )
                            # Scroll into view to avoid overlap by header/footer
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", port_dropdown)
                            time.sleep(0.5)
                        
                            # Wait until overlay or header disappears if exists
                            try:
                                WebDriverWait(self.driver, 5).until_not(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.userDetails"))
                                )
                            except TimeoutException:
                                pass  # continue anyway
                        
                            # Try clicking normally first
                            try:
                                port_dropdown.click()
                            except:
                                # If still intercepted, use JS click fallback
                                self.driver.execute_script("arguments[0].click();", port_dropdown)
                        
                            print("port dropdown opened")
                            time.sleep(1)
                        
                            search_input = wait.until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                            )
                            search_input.clear()
                            search_input.send_keys(location_code)
                            time.sleep(1.5)
                        
                            first_option = wait.until(
                                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                            )
                            self.driver.execute_script("arguments[0].click();", first_option)
                            print(f" Selected location: {location_code}")
                        
                        except Exception as e:
                            print(f" Port selection failed for {location_code}: {e}")
                            self.driver.refresh()
                            # break
                            continue

                        # --- Click Search ---
                        try:
                            time.sleep(1)
                            search_button = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Search']"))
                            )
                            # search_button.click()
                            self.driver.execute_script("arguments[0].click();", search_button)
                            print("🔍 Clicked Search button")
                            time.sleep(3)
                            try:
                                # Wait for popup to appear (if it does)
                                popup=WebDriverWait(self.driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                                )
                                message_text = popup.text.strip()
                                print(f"⚠️ Popup message detected: {message_text}")
                                print("⚠️ Popup detected — waiting for it to close...")
                            
                                # Wait for popup to disappear
                                popup=WebDriverWait(self.driver, 5).until_not(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                                )
                                print("✅ Popup closed automatically")
                                if message_text=="must not be empty" or message_text== "auth error..!!!" :
                                    self.driver.get("https://www.icegate.gov.in")
                                    WebDriverWait(self.driver, 100).until(
                                        EC.element_to_be_clickable((By.LINK_TEXT, "Login / Sign Up"))).click()
                                    print("click to login")
                        
                                    # Wait briefly to ensure the new tab is opened
                                    time.sleep(4)
                                    self.driver.refresh()
                                    # Get all window handles
                                    tabs = self.driver.window_handles
                                    
                                    # Switch to the newest tab (usually the last in the list)
                                    self.driver.switch_to.window(tabs[-1])
                                    
                                    # Now you're in the new tab; continue your automation
                                    print("Switched to new tab:", self.driver.current_url)
                        
                                    # self.driver.execute_script("window.scrollBy(0, 50);")
                                    # self.driver.refresh()
                        
                                    WebDriverWait(self.driver, 100).until(
                                        EC.presence_of_element_located((By.ID, "icegateId")))
                                    print("wait for icegateID")
                                    
                                    self.driver.find_element(By.ID, "icegateId").send_keys(self.username)
                                    self.driver.find_element(By.ID, "password").send_keys(self.password)
                                    original_url = self.driver.current_url
                                    WebDriverWait(self.driver, 100).until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LOGIN']"))).click()
                                    
                        
                                    # Wait until the URL changes (max 5 minutes)
                                    WebDriverWait(self.driver, 100).until(EC.url_changes(original_url))
                        
                                    print("change URL")
                                    time.sleep(1)
                                    self.driver.refresh()
                                    print("refresh page")
                                    time.sleep(2)

                                    # //*[@id="mat-expansion-panel-header-5"]/span[1]/h5
                                    # //*[@id="mat-expansion-panel-header-5"]/span[2]
                                    # //*[@id="mat-expansion-panel-header-5"]
                                    # services = self.driver.find_element(By.XPATH, "//h5[normalize-space()='Services']")
                                    try:
                                       time.sleep(1)
                                       services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
                                       print("services")
                                       self.driver.execute_script("arguments[0].click();", services)
                                       print("click on services")
                                    except:
                                        print("not work 5")
                                        services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                                        print("services")
                                        self.driver.execute_script("arguments[0].click();", services)
                                        print("click on services")
                                    
                        
                                    self.driver.execute_script("window.scrollBy(0, 100);")
                                    time.sleep(1)
                        
                                    
                        
                                    wait = WebDriverWait(self.driver, 100)
                        
                                    # Click the chevron corresponding to the "E-Payment" node
                                    # Enquiries = wait.until(EC.element_to_be_clickable((
                                    #     By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button/span/mat-icon'
                                    # )))
                                    Enquiries = wait.until(EC.element_to_be_clickable((
                                        By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
                                        
                                    )))
                                    Enquiries.click()
                                    print("click on Enquiries")
                    
                                    icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                                        By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
                                    )))
                                    icegate_enquiry_service.click()
                                    print("click on Icegate Enquiry services")
                    
                                    drawback_disbursement_status = wait.until(EC.element_to_be_clickable((
                                        By.XPATH, '/html/body/app-root/app-layout/div/div[2]/app-enquiries-layout/app-available-enquiries/div/div/div[1]/div/div[7]/div/div/span'
                                    )))
                                    drawback_disbursement_status.click()
                                    print("click on drawback disbursement status")
                                    continue
                                self.driver.refresh()
                                time.sleep(1)
                                current_from = current_to + timedelta(days=1)
                                time.sleep(1)
                                continue


                            
                            except TimeoutException:
                                print("✅ No popup appeared, continuing...")
                        except Exception as e:
                            print(f"⚠️ Search click failed: {e}")
                            break
                       

                        try:
                            time.sleep(2)
                            download_button = WebDriverWait(self.driver, 20).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Download Excel']"))
                            )
                            time.sleep(1)
                            # download_button.click()
                            self.driver.execute_script("arguments[0].click();", download_button)
                            time.sleep(3)
                            print("🔍 Clicked download button")
                            time.sleep(1)
                        except Exception as e:
                            print(f"data not shown")
                            self.driver.refresh()
                            time.sleep(1)
                            continue
                           
                        
                        time.sleep(7)

                        self.driver.refresh()
                        time.sleep(1)

                        current_from = current_to + timedelta(days=1)
                
                    print(f"✅ Finished all date ranges for port: {location_code}")
                
                print("\n🎯 All ports processed successfully.")
                return {"success": True, "message": "All ports processed successfully"}
            else:
                time.sleep(2)
    
            # return {"success": True, "message": f"Row {self.current_index + 1} processed successfully"}
            print(f"Row {self.current_index + 1} processed successfully")
        


                                # /*-------  Drawback Pendency Data -----*/

            # self.driver.refresh()
            time.sleep(2)
            print("wait for home")
            # winsound.Beep(1000, 1000)  
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "(//*[name()='svg'][@role='img'])[1]"))).click()
            print("click home")
           
            # WebDriverWait(self.driver, 100).until(
            #     EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Services']"))).click()
            # print("click services")
            try:
               time.sleep(1)
               services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
               print("services")
               self.driver.execute_script("arguments[0].click();", services)
               print("click on services")
            except:
                print("not work 5")
                services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                print("services")
                self.driver.execute_script("arguments[0].click();", services)
                print("click on services")
            

            self.driver.execute_script("window.scrollBy(0, 100);")
            time.sleep(1)

            wait = WebDriverWait(self.driver, 100)

            # Click the chevron corresponding to the "E-Payment" node
            # Enquiries = wait.until(EC.element_to_be_clickable((
            #     By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button/span/mat-icon'
            # )))
            Enquiries = wait.until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
            )))
            Enquiries.click()
            print("click on Enquiries")
            icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
            )))
            icegate_enquiry_service.click()
            print("click on Icegate Enquiry services")
            Drawback_enquiry = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[normalize-space()='Drawback Enquiry']"
            )))
            Drawback_enquiry.click()
            print("click on drawback Drawback_enquiry")
            time.sleep(2)

            global_start_date = None
            global_end_date = None
            
            for idx, row in enumerate(self.excel_data):
                location_code = str(row.get("Pendency Port", "")).strip()
                print(f"\n📍 Processing Port {idx+1}: {location_code}")
            
                def parse_date_safe(date_value):
                    if pd.isna(date_value):
                        return None
                    if isinstance(date_value, str):
                        try:
                            return datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S")
                        except ValueError:
                            try:
                                return datetime.strptime(date_value, "%Y-%m-%d")
                            except:
                                return None
                    return date_value if isinstance(date_value, datetime) else None
            
                # --- For first port, read dates from Excel ---
                if idx == 0:
                    start_date_obj = parse_date_safe(row.get("Start Date Pendency"))
                    end_date_obj = parse_date_safe(row.get("End Date Pendency"))
            
                    if not start_date_obj or not end_date_obj:
                        print(f"⚠️ Invalid or missing date(s) for first port: {location_code}")
                        break
            
                    # Save for reuse
                    global_start_date = start_date_obj
                    global_end_date = end_date_obj
                    print(f" Base Date Range Set: {global_start_date.strftime('%d-%m-%Y')} → {global_end_date.strftime('%d-%m-%Y')}")
            
                # --- For later ports, reuse first port’s dates ---
                else:
                    start_date_obj = global_start_date
                    end_date_obj = global_end_date
            
                current_from = start_date_obj
            
                # === Loop through all 30-day windows ===
                while current_from <= end_date_obj:
                    
                    current_to = current_from + timedelta(days=29)
                    if current_to > end_date_obj:
                        current_to = end_date_obj
            
                    from_str = current_from.strftime("%d-%m-%Y")
                    to_str = current_to.strftime("%d-%m-%Y")
                    time.sleep(1)
            
                    # --- Fill 'Date From' ---
                    self.driver.execute_script("""
                        var el = document.querySelector('input[formcontrolname="fromDate"]');
                        if (el) {
                            el.removeAttribute('readonly');
                            el.value = arguments[0];
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                        }
                    """, from_str) 
            
                    # --- Fill 'To' date ---
                    self.driver.execute_script("""
                        var el = document.querySelector('input[formcontrolname="toDate"]');
                        if (el) {
                            el.removeAttribute('readonly');
                            el.value = arguments[0];
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                        }
                    """, to_str)
                    
            
                    print(f"Date Range: {from_str} → {to_str}")
                    time.sleep(2)
            
                    # --- Select Port ---
                    # time.sleep(2)
                    # try:
                    #     port_dropdown = wait.until(
                    #         EC.element_to_be_clickable((By.XPATH, "//ng-select[contains(@formcontrolname,'locationCode') or contains(.,'Select Port')]"))
                    #     )
                    #     port_dropdown.click()
                    #     time.sleep(1)
                    #     print("port dropdown")
                    #     search_input = wait.until(
                    #         EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                    #     )
                    #     search_input.clear()
                    #     search_input.send_keys(location_code)
                    #     time.sleep(1.5)
            
                    #     first_option = wait.until(
                    #         EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                    #     )
                    #     self.driver.execute_script("arguments[0].click();", first_option)
                    #     print(f" Selected location: {location_code}")
                    # except Exception as e:
                    #     print(f" Port selection failed for {location_code}: {e}")
                    #     break

                    try:
                        # Wait for dropdown to be clickable in DOM
                        port_dropdown = wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//ng-select[contains(@formcontrolname,'locationCode') or contains(.,'Select Port')]")
                            )
                        )
                        # Scroll into view to avoid overlap by header/footer
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", port_dropdown)
                        time.sleep(0.5)
                    
                        # Wait until overlay or header disappears if exists
                        try:
                            WebDriverWait(self.driver, 5).until_not(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "div.userDetails"))
                            )
                        except TimeoutException:
                            pass  # continue anyway
                    
                        # Try clicking normally first
                        try:
                            port_dropdown.click()
                        except:
                            # If still intercepted, use JS click fallback
                            self.driver.execute_script("arguments[0].click();", port_dropdown)
                    
                        print("port dropdown opened")
                        time.sleep(1)
                    
                        search_input = wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                        )
                        search_input.clear()
                        search_input.send_keys(location_code)
                        time.sleep(1.5)
                    
                        first_option = wait.until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                        )
                        self.driver.execute_script("arguments[0].click();", first_option)
                        print(f" Selected location: {location_code}")
                    
                    except Exception as e:
                        print(f" Port selection failed for {location_code}: {e}")
                        self.driver.refresh()
                        continue
                        # break
            
                    # --- Click Search drawback pending status ---
                    try:
                        drawback_pending_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Drawback Pending Status']"))
                        )
                        drawback_pending_button.click()
                        print(" Clicked drawback pending status button")
                        time.sleep(3)
                      
                        try:
                            # Wait for popup to appear (if it does)
                            popup=WebDriverWait(self.driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                            )
                            message_text = popup.text.strip()
                            print(f"⚠️ Popup message detected: {message_text}")
                            print("⚠️ Popup detected — waiting for it to close...")
                        
                            # Wait for popup to disappear
                            popup=WebDriverWait(self.driver, 5).until_not(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                            )
                            print("✅ Popup closed automatically")

                            if message_text=="auth error..!!!" or message_text=="Service unavailable, Please try again after sometime":
                                self.driver.get("https://www.icegate.gov.in")
                                WebDriverWait(self.driver, 100).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, "Login / Sign Up"))).click()
                                print("click to login")
                    
                                # Wait briefly to ensure the new tab is opened
                                time.sleep(4)
                                self.driver.refresh()
                                # Get all window handles
                                tabs = self.driver.window_handles
                                
                                # Switch to the newest tab (usually the last in the list)
                                self.driver.switch_to.window(tabs[-1])
                                
                                # Now you're in the new tab; continue your automation
                                print("Switched to new tab:", self.driver.current_url)
                    
                                # self.driver.execute_script("window.scrollBy(0, 50);")
                                # self.driver.refresh()
                    
                                WebDriverWait(self.driver, 100).until(
                                    EC.presence_of_element_located((By.ID, "icegateId")))
                                print("wait for icegateID")
                                
                                self.driver.find_element(By.ID, "icegateId").send_keys(self.username)
                                self.driver.find_element(By.ID, "password").send_keys(self.password)
                                original_url = self.driver.current_url
                                WebDriverWait(self.driver, 100).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LOGIN']"))).click()
                                # original_url = self.driver.current_url
                                time.sleep(10)
                    
                                # Wait until the URL changes (max 5 minutes)
                                WebDriverWait(self.driver, 100).until(EC.url_changes(original_url))
                    
                                print("change URL")
                                time.sleep(1)
                                

                                
                                # WebDriverWait(self.driver, 100).until(
                                #     EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Services']"))).click()
                                # print("click services")
                                try:
                                   time.sleep(1)
                                   services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
                                   print("services")
                                   self.driver.execute_script("arguments[0].click();", services)
                                   print("click on services")
                                   self.driver.execute_script("window.scrollBy(0,50)")
                                except:
                                    print("not work 5")
                                    services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                                    print("services")
                                    self.driver.execute_script("arguments[0].click();", services)
                                    print("click on services")
                                    self.driver.execute_script("window.scrollBy(0,50)")
                                

                    
                                self.driver.execute_script("window.scrollBy(0, 100);")
                                time.sleep(1)
                    
                                wait = WebDriverWait(self.driver, 100)
                    
                                # Click the chevron corresponding to the "E-Payment" node
                                Enquiries = wait.until(EC.element_to_be_clickable((
                                    By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
                                )))
                                Enquiries.click()
                                print("click on Enquiries")
                                icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                                    By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
                                )))
                                icegate_enquiry_service.click()
                                print("click on Icegate Enquiry services")
                                Drawback_enquiry = wait.until(EC.element_to_be_clickable((
                                    By.XPATH, "//span[normalize-space()='Drawback Enquiry']"
                                )))
                                Drawback_enquiry.click()
                                print("click on drawback Drawback_enquiry")
                                time.sleep(2)
                                continue
                            self.driver.refresh()
                            time.sleep(1)
                            current_from = current_to + timedelta(days=1)
                            
                            continue
                        
                        except TimeoutException:
                            print("✅ No popup appeared, continuing...")
                    except Exception as e:
                        print(f" drawback pending status click failed: {e}")
                        break
                    
                    # time.sleep(5)
                    # download_button = wait.until(
                    #     EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Download Excel']"))
                    # )
                    # download_button.click()
                    # print(" Clicked download button")
                    try:
                        download_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Download Excel']"))
                        )
                        time.sleep(1)
                        download_button.click()
                        print(" Clicked download button")
                        time.sleep(1)
                        self.driver.refresh()
                        
                        
                    except Exception as e:
                        print(f"data not shown")
                        drawback_pending_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Drawback Pending Status']"))
                        )
                        drawback_pending_button.click()
                        print(" Clicked drawback pending button")
                        time.sleep(1)
                    
                    time.sleep(2)
                    self.driver.refresh()
                    time.sleep(2)
                    current_from = current_to + timedelta(days=1)
            
                print(f" Finished all date ranges for port: {location_code}")
            
            print("\n All ports processed successfully.")


            # /*------- GST realisation and Credit Data -----*
            # self.driver.refresh()
            time.sleep(2)
            print("wait for home")
            # winsound.Beep(1000, 1000)  
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "(//*[name()='svg'][@role='img'])[1]"))).click()
            print("click home")
            
            # WebDriverWait(self.driver, 100).until(
            #     EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Services']"))).click()
            # print("click services")

            try:
               time.sleep(1)
               services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
               print("services")
               self.driver.execute_script("arguments[0].click();", services)
               print("click on services")
               self.driver.execute_script("window.scrollBy(0,50)")
            except:
                print("not work 5")
                services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                print("services")
                self.driver.execute_script("arguments[0].click();", services)
                print("click on services")
                self.driver.execute_script("window.scrollBy(0,50)")
            

            self.driver.execute_script("window.scrollBy(0, 100);")
            time.sleep(1)

            wait = WebDriverWait(self.driver, 100)

            # Click the chevron corresponding to the "E-Payment" node
            Enquiries = wait.until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
            )))
            Enquiries.click()
            print("click on Enquiries")
            icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
            )))
            icegate_enquiry_service.click()
            print("click on Icegate Enquiry services")
            IGST_Scroll_status = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[normalize-space()='IGST Scroll Sanctioned Status']"
            )))
            IGST_Scroll_status.click()
            print("click on IGST Scroll status")
            time.sleep(2)

             # ---- Excel Output Setup ----
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            OUTPUT_FILE = os.path.join(
                downloads_path,
                f"IGST_Scroll_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
    
            final_data = []
            final_headers = []
    
            global_start_date = None
            global_end_date = None
    
            # ---- Loop through Ports ----
            for idx, row in enumerate(self.excel_data):
                location_code = str(row.get("IGST Port", "")).strip()
                print(f"\n Processing Port {idx + 1}: {location_code}")
    
                def parse_date_safe(date_value):
                    if pd.isna(date_value):
                        return None
                    if isinstance(date_value, str):
                        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                            try:
                                return datetime.strptime(date_value, fmt)
                            except ValueError:
                                pass
                        return None
                    return date_value if isinstance(date_value, datetime) else None
    
                if idx == 0:
                    start_date_obj = parse_date_safe(row.get("Start Date IGST"))
                    end_date_obj = parse_date_safe(row.get("End Date IGST"))
                    if not start_date_obj or not end_date_obj:
                        print(f" Invalid or missing date(s) for first port: {location_code}")
                        break
                    global_start_date = start_date_obj
                    global_end_date = end_date_obj
                    print(f" Base Date Range Set: {global_start_date.strftime('%d-%m-%Y')} → {global_end_date.strftime('%d-%m-%Y')}")
                else:
                    start_date_obj = global_start_date
                    end_date_obj = global_end_date
    
                current_from = start_date_obj
    
                # ---- Loop through date ranges ----
                while current_from <= end_date_obj:
                    location_code = str(row.get("IGST Port", "")).strip()
                     # ⬅️ Goes directly to next while iteration
                    if location_code in [None, "", "nan"]:
                        print("⚠️ Skipping this iteration — Location Code is empty or NaN")
                        current_from = current_to + timedelta(days=1)
                        continue
                    else: 
                        current_to = current_from + timedelta(days=5)
                        if current_to > end_date_obj:
                            current_to = end_date_obj
        
                        from_str = current_from.strftime("%d-%m-%Y")
                        to_str = current_to.strftime("%d-%m-%Y")
        
                        # Fill 'From' date
                        self.driver.execute_script("""
                            var el = document.querySelector('input[formcontrolname="startDate"]');
                            if (el) {
                                el.removeAttribute('readonly');
                                el.value = arguments[0];
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                            }
                        """, from_str)
        
                        # Fill 'To' date
                        self.driver.execute_script("""
                            var el = document.querySelector('input[formcontrolname="endDate"]');
                            if (el) {
                                el.removeAttribute('readonly');
                                el.value = arguments[0];
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                            }
                        """, to_str)
        
                        print(f"Date Range: {from_str} → {to_str}")
        
                        # ---- Select Port ----
                        # try:
                        #     time.sleep(2)
                        #     port_dropdown = wait.until(
                        #         EC.element_to_be_clickable((By.XPATH, "//ng-select[contains(@formcontrolname,'locationCode')]"))
                        #     )
                        #     port_dropdown.click()
                        #     search_input = wait.until(
                        #         EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                        #     )
                        #     search_input.clear()
                        #     search_input.send_keys(location_code)
                        #     time.sleep(1.5)
                        #     first_option = wait.until(
                        #         EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                        #     )
                        #     self.driver.execute_script("arguments[0].click();", first_option)
                        #     print(f" Selected location: {location_code}")
                        # except Exception as e:
                        #     print(f" Port selection failed for {location_code}: {e}")
                        #     break

                        try:
                            # Wait for dropdown to be clickable in DOM
                            port_dropdown = wait.until(
                                EC.presence_of_element_located(
                                    (By.XPATH, "//ng-select[contains(@formcontrolname,'locationCode') or contains(.,'Select Port')]")
                                )
                            )
                            # Scroll into view to avoid overlap by header/footer
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", port_dropdown)
                            time.sleep(0.5)
                        
                            # Wait until overlay or header disappears if exists
                            try:
                                WebDriverWait(self.driver, 5).until_not(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.userDetails"))
                                )
                            except TimeoutException:
                                pass  # continue anyway
                        
                            # Try clicking normally first
                            try:
                                port_dropdown.click()
                            except:
                                # If still intercepted, use JS click fallback
                                self.driver.execute_script("arguments[0].click();", port_dropdown)
                        
                            print("port dropdown opened")
                            time.sleep(1)
                        
                            search_input = wait.until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='combobox'] input[type='text']"))
                            )
                            search_input.clear()
                            search_input.send_keys(location_code)
                            time.sleep(1.5)
                        
                            first_option = wait.until(
                                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ng-option"))
                            )
                            self.driver.execute_script("arguments[0].click();", first_option)
                            print(f" Selected location: {location_code}")
                        
                        except Exception as e:
                            print(f" Port selection failed for {location_code}: {e}")
                            self.driver.refresh()
                            continue    
                            # break
        
                        # ---- Click Search ----
                        try:
                            search_button = wait.until(EC.element_to_be_clickable(
                                (By.XPATH, "//button[normalize-space()='Search']")))
                            search_button.click()
                            print(" Clicked Search button")
                            time.sleep(2)
                            try:
                                # Wait for popup to appear (if it does)
                                popup=WebDriverWait(self.driver, 3).until(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                                )
                                message_text = popup.text.strip()
                                print(f"⚠️ Popup message detected: {message_text}")
                                print("⚠️ Popup detected — waiting for it to close...")
                            
                                # Wait for popup to disappear
                                popup=WebDriverWait(self.driver, 5).until_not(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'popup') or contains(@role, 'alert') or contains(text(),'Success') or contains(text(),'Error')]"))
                                )
                                print("✅ Popup closed automatically")
    
                                if message_text=="auth error..!!!" or message_text=="must not be empty":
                                    self.driver.get("https://www.icegate.gov.in")
                                    WebDriverWait(self.driver, 100).until(
                                        EC.element_to_be_clickable((By.LINK_TEXT, "Login / Sign Up"))).click()
                                    print("click to login")
                        
                                    # Wait briefly to ensure the new tab is opened
                                    time.sleep(4)
                                    self.driver.refresh()
                                    # Get all window handles
                                    tabs = self.driver.window_handles
                                    
                                    # Switch to the newest tab (usually the last in the list)
                                    self.driver.switch_to.window(tabs[-1])
                                    
                                    # Now you're in the new tab; continue your automation
                                    print("Switched to new tab:", self.driver.current_url)
                        
                                    # self.driver.execute_script("window.scrollBy(0, 50);")
                                    # self.driver.refresh()
                        
                                    WebDriverWait(self.driver, 100).until(
                                        EC.presence_of_element_located((By.ID, "icegateId")))
                                    print("wait for icegateID")
                                    
                                    self.driver.find_element(By.ID, "icegateId").send_keys(self.username)
                                    self.driver.find_element(By.ID, "password").send_keys(self.password)
                                    original_url = self.driver.current_url
                                    WebDriverWait(self.driver, 100).until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LOGIN']"))).click()
                                    
                        
                                    # Wait until the URL changes (max 5 minutes)
                                    WebDriverWait(self.driver, 100).until(EC.url_changes(original_url))
                        
                                    print("change URL")
                                    time.sleep(1)
                                    
    
                                    
                                    # WebDriverWait(self.driver, 100).until(
                                    #     EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Services']"))).click()
                                    # print("click services")

                                    try:
                                       time.sleep(1)
                                       services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-5"]')
                                       print("services")
                                       self.driver.execute_script("arguments[0].click();", services)
                                       print("click on services")
                                    except:
                                        print("not work 5")
                                        services = self.driver.find_element(By.XPATH, '//*[@id="mat-expansion-panel-header-24"]')
                                        print("services")
                                        self.driver.execute_script("arguments[0].click();", services)
                                        print("click on services")
                                    
                        
                                    self.driver.execute_script("window.scrollBy(0, 100);")
                                    time.sleep(1)
                        
                                    wait = WebDriverWait(self.driver, 100)
                        
                                    # Click the chevron corresponding to the "E-Payment" node
                                    try:

                                        Enquiries = wait.until(EC.element_to_be_clickable((
                                            By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button'
                                        )))
                                        # //*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button/span/mat-icon
                                        # //*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button
                                        Enquiries.click()
                                        print("click on Enquiries")
                                    except:
                                        Enquiries = wait.until(EC.element_to_be_clickable((
                                            By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button/span/mat-icon'
                                        )))
                                        
                                        # //*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/div/button
                                        Enquiries.click()
                                        print("click on Enquiries")

                                    icegate_enquiry_service = wait.until(EC.element_to_be_clickable((
                                        By.XPATH, '//*[@id="cdk-accordion-child-5"]/div/div/mat-tree/mat-nested-tree-node[4]/li/ul/mat-tree-node/li/a'
                                    )))
                                    icegate_enquiry_service.click()
                                    print("click on Icegate Enquiry services")
                                    IGST_Scroll_status = wait.until(EC.element_to_be_clickable((
                                        By.XPATH, "//span[normalize-space()='IGST Scroll Sanctioned Status']"
                                    )))
                                    IGST_Scroll_status.click()
                                    print("click on IGST Scroll status")
                                    time.sleep(2)
                                    continue
                                self.driver.refresh()
                                time.sleep(1)
                                current_from = current_to + timedelta(days=1)
                                
                                continue
                            
                            except TimeoutException:
                                print("✅ No popup appeared, continuing...")
                        except Exception as e:
                            print(f" Search click failed: {e}")
                            # Move to next date range
                            current_from = current_to + timedelta(days=1)
                            continue  
                        
                        # ---- Wait for Table or No Records ----
                        try:
                            # Wait for either table or "No records found"
                            WebDriverWait(self.driver, 25).until(
                                EC.any_of(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "div#tablerecords table")),
                                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'No records found')]"))
                                )
                            )
                        
                            # --- Check if "No records found" is visible ---
                            no_record = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'No records found')]")
                            if no_record:
                                print(" No records found for this date range — moving to next date range.\n")
                                # Increment to next range before continuing
                                current_from = current_to + timedelta(days=1)
                                continue  
                        
                            print("Table loaded for this date range")
                        
                            # ---- Select 100 items per page ----
                            try:
                                dropdown = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-0")))
                                dropdown.click()
                                panel = wait.until(EC.presence_of_element_located(
                                    (By.XPATH, "//div[contains(@class,'mat-select-panel')]")))
                                option_100 = wait.until(EC.element_to_be_clickable((
                                    By.XPATH, "//span[contains(@class,'mat-option-text') and normalize-space(text())='100']")))
                                option_100.click()
                                time.sleep(2)
                                print(" Set 100 items per page")
                            except Exception:
                                print(" Could not set 100 items per page, continuing anyway...")
                        
                        except TimeoutException:
                            print(" No results or table appeared — skipping this date range.\n")
                            current_from = current_to + timedelta(days=1)
                            continue  # move to next date range
                        
                        # ---- Extract table with pagination ----
                        while True:
                            time.sleep(2)
                            table = self.driver.find_element(By.CSS_SELECTOR, "div#tablerecords table")
                        
                            if not final_headers:
                                header_elements = table.find_elements(By.TAG_NAME, "th")
                                final_headers = [h.text.strip() for h in header_elements] + ["Port Code"]
                        
                            rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
                            for r in rows:
                                cols = [td.text.strip() for td in r.find_elements(By.TAG_NAME, "td")]
                                if any(cols):
                                    cols.append(location_code)
                                    final_data.append(cols)
                        
                            try:
                                next_button = self.driver.find_element(By.CSS_SELECTOR, "button.mat-paginator-navigation-next")
                                disabled = next_button.get_attribute("disabled")
                                if disabled:
                                    print(" Last page reached.")
                                    break
                                else:
                                    next_button.click()
                            except Exception:
                                print(" No next button. Assuming last page.")
                                break
                            self.driver.execute_script("window.scrollTo(0, 0);")
                            time.sleep(2)
                       
                        # Go to next date window
                        self.driver.refresh()
                        time.sleep(3)
                        current_from = current_to + timedelta(days=1)
    
                print(f" Completed port: {location_code}")
    
            # ---- After all ports ----
            if final_data:
                df = pd.DataFrame(final_data, columns=final_headers)
                df.to_excel(OUTPUT_FILE, index=False)
                print(f"\n🎯 All port data saved to: {OUTPUT_FILE}")
                # os.startfile(os.path.dirname(OUTPUT_FILE))
            else:
                print("⚠️ No data found to export.")


            # /*------- RODTEP Scroll Data -----*/

            self.driver.refresh()
            time.sleep(2)
            self.driver.refresh()
            print("wait for home")
            print("wait")
            time.sleep(2)
            # winsound.Beep(1000, 1000)  
            try:
                WebDriverWait(self.driver, 100).until(
                    EC.element_to_be_clickable((By.XPATH, "//ul[@class='navigation relative-nav']//a[@title='Home']//*[name()='svg']//*[name()='path' and contains(@d,'M280.37 14')]"))).click()
                    
                print("click home")
            except:
                
                WebDriverWait(self.driver, 100).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="details"]/div/div[1]/div/nav/div[1]/ul/li[1]/a/svg'))).click()
                    
                print("click home")
            
            
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='eScrip - Credit Ledger']"))).click()
            print("click eScrip - credit ledger")

            self.driver.execute_script("window.scrollBy(0, 100);")
            time.sleep(1) 
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//h2[@id='schemeCount']"))).click()
            print("click eScrip - credit ledger")
            self.driver.execute_script("window.scrollBy(0, 200);")
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='View Complete Details']"))).click()
            print("click view complete details")
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Scrip Generation')]"))).click()
            print("click Scrip generation")
            
               # ---------- 3️⃣ Wait until the table or dropdown appears ----------
            wait = WebDriverWait(self.driver, 60)
            wait.until(EC.presence_of_element_located((By.XPATH, "//mat-select[@aria-label='Items per page:']")))
            
            # ---------- 4️⃣ Click the dropdown properly ----------
            try:
                dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select[@aria-label='Items per page:']")))
                self.driver.execute_script("arguments[0].click();", dropdown)
                time.sleep(1)
            
                # Wait for overlay container to appear (Angular Material adds class 'mat-select-panel')
                panel = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'mat-select-panel')]")))
            
                # Click the 100 option from overlay
                option_100 = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//mat-option//span[normalize-space()='100']"))
                )
                self.driver.execute_script("arguments[0].click();", option_100)
                print("✅ Selected 100 items per page.")
                time.sleep(3)
            except Exception as e:
                print("⚠️ Unable to select 100 items:", e)
            
            # ---------- 5️⃣ Extract table data ----------
            all_data = []
            page = 1
            
            while True:
                print(f"\n📄 Extracting Page {page}...")
            
                # Wait for table rows
                rows = wait.until(
                    EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class,'mat-table')]/tbody/tr"))
                )
            
                for row in rows:
                    cols = row.find_elements(By.XPATH, ".//td")
                    row_data = [col.text.strip() for col in cols]
                    if any(row_data):
                        all_data.append(row_data)
            
                print(f"✅ Page {page} data extracted ({len(rows)} rows).")
            
                # Check if "Next page" is enabled
                try:
                    next_button = self.driver.find_element(By.XPATH, "//button[contains(@aria-label,'Next page')]")
                    if next_button.get_attribute("disabled"):
                        print("📘 No more pages.")
                        break
                    else:
                        self.driver.execute_script("arguments[0].click();", next_button)
                        page += 1
                        time.sleep(3)
                except Exception:
                    print("⚠️ No Next button found.")
                    break
            
            # ---------- 6️⃣ Save to Excel ----------
            if all_data:
                headers = [h.text.strip() for h in self.driver.find_elements(By.XPATH, "//table[contains(@class,'mat-table')]/thead/tr/th")]
                df = pd.DataFrame(all_data, columns=headers if headers else None)
            
                download_path = os.path.join(os.path.expanduser("~"), "Downloads", "RODTEP_Scroll_data.xlsx")
                df.to_excel(download_path, index=False)
                print(f"\n💾 Excel saved at: {download_path}")
            else:
                print("⚠️ No data found to export.")
    
            
            # /*------- RODTEP Scrip Data -----*
            # self.driver.refresh()
            time.sleep(2)
            print("wait for home")
            # winsound.Beep(1000, 1000)  
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "(//*[name()='svg'][@role='img'])[1]"))).click()
            print("click home")
            
            WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='eScrip - Credit Ledger']"))).click()
            print("click eScrip - credit ledger")

            self.driver.execute_script("window.scrollBy(0, 100);")
            time.sleep(1) 
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//h2[@id='schemeCount']"))).click()
            print("click eScrip - credit ledger")
            self.driver.execute_script("window.scrollBy(0, 200);")
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='View Complete Details']"))).click()
            print("click view complete details")
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Scrip Details')]"))).click()
            print("click Scrip Details")
            time.sleep(2)
            self.driver.execute_script("window.scrollBy(0, 150);")
            # -----------------------------
            # Wait for main table to load
            # -----------------------------
            wait = WebDriverWait(self.driver, 30)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.mat-table")))
            
            data = []
            
            # -----------------------------------
            # Loop through all table pages
            # -----------------------------------
            while True:
                time.sleep(1)
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table.mat-table tbody tr")
            
                for i, row in enumerate(rows, start=1):
                    cols = row.find_elements(By.CSS_SELECTOR, "td")
                    if len(cols) < 9:
                        continue
            
                    scrip_no = cols[0].text.strip()
                    scrip_issue_date = cols[1].text.strip()
                    scrip_exp_date = cols[2].text.strip()
                    scrip_issued_amount = cols[3].text.strip()
                    scrip_balance = cols[4].text.strip()
                    scrip_transfer_date = cols[5].text.strip()
                    scrip_status = cols[6].text.strip()
            
                    scroll_numbers = []
                    sb_numbers = []
            
                    # ===============================
                    # Extract Scroll Numbers
                    # ===============================
                    try:
                        time.sleep(1.2)
                        scroll_link = cols[7].find_element(By.TAG_NAME, "a")
                        self.driver.execute_script("arguments[0].click();", scroll_link)
                        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-container")))
                        time.sleep(1)
            
                        try:
                            time.sleep(0.2)
                            scroll_container = wait.until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-container"))
                            )
                            scroll_items = scroll_container.find_elements(By.XPATH, ".//li")
                            scroll_numbers = [li.text.strip() for li in scroll_items if li.text.strip()]
                        except Exception:
                            scroll_numbers = []
            
                        # Close popup
                        try:
                            time.sleep(1)
                            close_btn = self.driver.find_element(
                                By.XPATH, "//span[@class='material-icons pointer']"
                            )
                            self.driver.execute_script("arguments[0].click();", close_btn)
                        except:
                            pass
            
                        # Clear overlay focus
                        time.sleep(0.7)
                        self.driver.execute_script("document.querySelector('body').click();")
                        time.sleep(0.7)
            
                    except Exception as e:
                        print(f"⚠️ Scroll popup error at row {i}: {e}")
                        scroll_numbers = []
            
                    # ===============================
                    # Extract SB Numbers
                    # ===============================
                    try:
                        time.sleep(1.2)
                        sb_link = cols[8].find_element(By.TAG_NAME, "a")
                        time.sleep(1)
                        self.driver.execute_script("arguments[0].click();", sb_link)
                        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-container")))
                        time.sleep(1)
            
                        try:
                            time.sleep(0.2)
                            sb_container = wait.until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-container"))
                            )
                            sb_items = sb_container.find_elements(By.XPATH, ".//li")
                            sb_numbers = [li.text.strip() for li in sb_items if li.text.strip()]
                        except Exception:
                            sb_numbers = []
            
                        # Close popup
                        try:
                            time.sleep(1)
                            close_btn = self.driver.find_element(
                                By.XPATH, "//span[@class='material-icons pointer']"
                            )
                            self.driver.execute_script("arguments[0].click();", close_btn)
                        except:
                            pass
            
                        # Clear overlay focus
                        time.sleep(0.7)
                        self.driver.execute_script("document.querySelector('body').click();")
                        time.sleep(0.7)
            
                    except Exception as e:
                        print(f"⚠️ SB popup error at row {i}: {e}")
                        sb_numbers = []
            
                    # ===============================
                    # Append Row Data
                    # ===============================
                    max_len = max(len(scroll_numbers), len(sb_numbers), 1)
                    for j in range(max_len):
                        scroll_val = scroll_numbers[j] if j < len(scroll_numbers) else ""
                        sb_val = sb_numbers[j] if j < len(sb_numbers) else ""
            
                        data.append({
                            "Scrip No": scrip_no,
                            "Scrip Issue Date": scrip_issue_date,
                            "Scrip Exp Date": scrip_exp_date,
                            "Scrip Issued Amount": scrip_issued_amount,
                            "Scrip Balance": scrip_balance,
                            "Scrip Transfer Date": scrip_transfer_date,
                            "Scrip Status": scrip_status,
                            "Scroll Number": scroll_val,
                            "SB Number": sb_val
                        })
            
                    print(f"✅ Row {i}: Scrolls {len(scroll_numbers)}, SBs {len(sb_numbers)}")
            
                # ===============================
                # Pagination
                # ===============================
                try:
                    time.sleep(1)
                    next_btn = self.driver.find_element(
                        By.XPATH, "//button[@aria-label='Next page' and not(@disabled)]"
                    )
                    self.driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(2)
                except Exception:
                    print("✅ No more pages. Extraction complete.")
                    break
            
            # ==========================================
            # ✅ Always save Excel after extraction
            # ==========================================
            try:
                if data:
                    df = pd.DataFrame(data)
                    # today = datetime.date.today().strftime("%Y-%m-%d")
                    excel_path = os.path.join(
                        os.path.expanduser("~"), "Downloads", f"ICEGATE_Scrip_Details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )
                    df.to_excel(excel_path, index=False)
                    print(f"📁 Excel saved successfully: {excel_path}")
                else:
                    print("⚠️ No data extracted, Excel not created.")
            except Exception as e:
                print(f"❌ Excel save error: {e}")
            return {"success": True, "message": "All ports processed successfully"}
        

        except Exception as e:
            return {"success": False, "message": f"Error in fill_certificate: {e}"}

    # def close_browser(self):
    #     if self.driver:
    #         self.driver.quit()
