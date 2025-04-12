import os
import pefile
import glob
import csv  # додаємо імпорт для роботи з csv.writer

# Шлях до файлів .exe
files = glob.glob('c:\\MalwareSamples\\*.exe')

if not files:
    print("Файли не знайдені у папці c:\\MalwareSamples\\")
else:
    # Відкриваємо CSV-файл
    with open('MalwareArtifacts.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        
        # Записуємо заголовок
        writer.writerow([
            "AddressOfEntryPoint", "MajorLinkerVersion", "MajorImageVersion",
            "MajorOperatingSystemVersion", "DllCharacteristics", "SizeOfStackReserve",
            "NumberOfSections", "ResourceSize"
        ])
        
        # Обробляємо кожен файл .exe
        for file in files:
            try:
                suspect_pe = pefile.PE(file)

                # Записуємо дані в CSV
                writer.writerow([
                    suspect_pe.OPTIONAL_HEADER.AddressOfEntryPoint,
                    suspect_pe.OPTIONAL_HEADER.MajorLinkerVersion,
                    suspect_pe.OPTIONAL_HEADER.MajorImageVersion,
                    suspect_pe.OPTIONAL_HEADER.MajorOperatingSystemVersion,
                    suspect_pe.OPTIONAL_HEADER.DllCharacteristics,
                    suspect_pe.OPTIONAL_HEADER.SizeOfStackReserve,
                    suspect_pe.FILE_HEADER.NumberOfSections,
                    suspect_pe.OPTIONAL_HEADER.DATA_DIRECTORY[2].Size
                ])
            except Exception as e:
                print(f"Помилка при обробці файлу {file}: {e}")