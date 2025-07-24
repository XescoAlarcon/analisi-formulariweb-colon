import os
import win32com.client
from collections import Counter
from datetime import datetime
import pandas as pd  # Añade esta importación al inicio del archivo

# Instalar, una única vez, los módulos requeridos en una terminal con derechos administrativos
# python -m pip install pywin32 pandas openpyxl

# Preguntar el año al usuario
anyo = input("Introduce el año (formato 4 dígitos, dejar vacío para todos): ").strip()
if anyo and not (anyo.isdigit() and len(anyo) == 4):
    print("Año no válido. Debe tener 4 dígitos o estar vacío.")
    exit(1)

# Ruta al archivo PST
pst_path = os.path.join(os.getcwd(), "colon.pst")

# Iniciar Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Abrir el archivo PST
outlook.AddStore(pst_path)

# Buscar el almacén recién añadido
for store in outlook.Stores:
    if store.FilePath.lower() == pst_path.lower():
        root_folder = store.GetRootFolder()
        print("\nCarpetas en colon.pst:")
        for folder in root_folder.Folders:
            print("-", folder.Name)
        # Buscar la carpeta "FORMULARI WEB"
        target_folder = None
        for folder in root_folder.Folders:
            if folder.Name.upper() == "FORMULARI WEB":
                target_folder = folder
                break
        if target_folder:
            print(f"\nProcesando e-mails en 'FORMULARI WEB' cuyo asunto empieza por 'Formulari còlon'" +
                  (f" y año {anyo}..." if anyo else " (todos los años)...") + ":")
            hombres = 0
            mujeres = 0
            datos_hombres = []
            datos_mujeres = []
            motivos = []
            detalle_motivos_otros_hombres = []
            detalle_motivos_otros_mujeres = []
            anyo_actual = datetime.now().year
            for item in target_folder.Items:
                if hasattr(item, "Subject") and hasattr(item, "ReceivedTime"):
                    if item.Subject.startswith("Formulari còlon") or item.Subject.startswith("Formulario colon"):
                        # Filtrar por año de ReceivedTime solo si se ha introducido
                        if not anyo or str(item.ReceivedTime.year) == anyo:
                            # Analizar el cuerpo del mensaje
                            if hasattr(item, "Body"):
                                lineas = item.Body.splitlines()
                                # Quedarnos con la primera línea para el Motivo
                                if len(lineas) > 0:
                                    motivo_linea = lineas[0]
                                    motivo= motivo_linea[len("Motiu de la consulta: "):].strip()
                                    motivo= motivo_linea[len("Motivo de la consulta: "):].strip()
                                    motivos.append(motivo)                                                                        
                                # Quedarnos con la cuarta línea que empieza por 'CIP:'
                                if len(lineas) >= 4:
                                    cip_linea = lineas[3]
                                    if len(cip_linea) > 0:
                                        try:
                                            # El décimo caracter de cip_linea (índice 9)
                                            cip_sexo = cip_linea[9]
                                            # Año de nacimiento: carácteres 11 y 12 (año)
                                            fecha_nacimiento = cip_linea[10:12]
                                            anyo_nacimiento = int('19' + fecha_nacimiento)
                                            edad = anyo_actual - anyo_nacimiento
                                            if edad >= 80:
                                                edad = -1
                                        except Exception:
                                            edad = None
                                        if cip_sexo == "0":
                                            hombres += 1
                                            if edad is not None:
                                                datos_hombres.append([edad, motivo])
                                                if (motivo == 'Altres consultes'):
                                                    otros_linea = lineas[6]
                                                    # El texto de otros aparece a partir del caracter 5
                                                    otros = otros_linea[5:]
                                                    detalle_motivos_otros_hombres.append(otros)
                                        elif cip_sexo == "1":
                                            mujeres += 1
                                            if edad is not None:
                                                datos_mujeres.append([edad, motivo])
                                                if (motivo == 'Altres consultes'):
                                                    otros_linea = lineas[6]
                                                    # El texto de otros aparece a partir del caracter 5
                                                    otros = otros_linea[5:]
                                                    detalle_motivos_otros_mujeres.append(otros)

            print(f"\n======================================")
            print(f"Total consultas:\t\t {hombres + mujeres}")
            print(f"======================================")
            
            print(f"\nTotal hombres:\t\t\t {hombres}")
            print(f"--------------------------------------")
            # Mostrar distribución de edades
            if datos_hombres:
                menores_60_h = sum(1 for e, _ in datos_hombres if e != -1 and e < 60)
                mayores_igual_60_h = sum(1 for e, _ in datos_hombres if e != -1 and e >= 60)
                missing_h = sum(1 for e, _ in datos_hombres if e == -1 )
                print(f"\tHombres < 60 años:\t {menores_60_h}")
                print(f"\tHombres >= 60 años:\t {mayores_igual_60_h}")
                print(f"\tHombres missing:\t {missing_h}")

                # Distribución por edad y motivo (hombres)
                print("\nDistribución por edad y motivo (hombres):")
                dist_hombres = {}
                for edad, motivo in datos_hombres:
                    if edad == -1:
                        grupo = ">=80 o edad desconocida"
                    elif edad < 60:
                        grupo = "<60"
                    else:
                        grupo = ">=60"
                    if grupo not in dist_hombres:
                        dist_hombres[grupo] = []
                    dist_hombres[grupo].append(motivo)
                for grupo, motivos in dist_hombres.items():
                    print(f"  Hombres {grupo} años ({len(motivos)}):")
                    motivos_count = Counter(motivos)
                    for motivo, count in motivos_count.items():
                        print(f"    - {motivo}: {count}")

            print(f"\nTotal mujeres:\t\t\t {mujeres} ")
            print(f"--------------------------------------")
            if datos_mujeres:
                menores_60_m = sum(1 for e, _ in datos_mujeres if e != -1 and e < 60)
                mayores_igual_60_m = sum(1 for e, _ in datos_mujeres if e != -1 and e >= 60)
                missing_m = sum(1 for e, _ in datos_mujeres if e == -1 )
                print(f"\tMujeres < 60 años:\t {menores_60_m}")
                print(f"\tMujeres >= 60 años:\t {mayores_igual_60_m}")
                print(f"\tMujeres missing:\t {missing_m}")

                # Distribución por edad y motivo (mujeres)
                print("\nDistribución por edad y motivo (mujeres):")
                dist_mujeres = {}
                for edad, motivo in datos_mujeres:
                    if edad == -1:
                        grupo = ">=80 o edad desconocida"
                    elif edad < 60:
                        grupo = "<60"
                    else:
                        grupo = ">=60"
                    if grupo not in dist_mujeres:
                        dist_mujeres[grupo] = []
                    dist_mujeres[grupo].append(motivo)
                for grupo, motivos in dist_mujeres.items():
                    print(f"  Mujeres {grupo} años ({len(motivos)}):")
                    motivos_count = Counter(motivos)
                    for motivo, count in motivos_count.items():
                        print(f"    - {motivo}: {count}")

            # Preguntar si se quiere ver el detalle de los otros motivos
            ver_detalle = input("\n¿Quieres ver el detalle de los 'otros motivos'? (s/n): ").strip().lower()
            if ver_detalle == "s":
                print("\nDetalle de 'otros motivos' (HOMBRES):")
                if detalle_motivos_otros_hombres:
                    print("-" * 40)
                    for i, motivo in enumerate(detalle_motivos_otros_hombres, 1):
                        print(f"{i}. {motivo}")
                    print("-" * 40)
                else:
                    print("No hay motivos adicionales para hombres.")

                print("\nDetalle de 'otros motivos' (MUJERES):")
                if detalle_motivos_otros_mujeres:
                    print("-" * 40)
                    for i, motivo in enumerate(detalle_motivos_otros_mujeres, 1):
                        print(f"{i}. {motivo}")
                    print("-" * 40)
                else:
                    print("No hay motivos adicionales para mujeres.")                        

            # Preguntar si se quiere exportar a Excel
            exportar_excel = input("\n¿Quieres exportar los datos a excel (s/n)? ").strip().lower()
            if exportar_excel == "s":
                # Crear DataFrames para hombres y mujeres
                df_hombres = pd.DataFrame(datos_hombres, columns=["Edad", "Motivo"])
                df_hombres["Sexo"] = "Hombre"
                df_mujeres = pd.DataFrame(datos_mujeres, columns=["Edad", "Motivo"])
                df_mujeres["Sexo"] = "Mujer"
                df = pd.concat([df_hombres, df_mujeres], ignore_index=True)

                # Añadir detalles de "otros motivos"
                otros_hombres = pd.DataFrame({
                    "Sexo": ["Hombre"] * len(detalle_motivos_otros_hombres),
                    "Detalle_Otros": detalle_motivos_otros_hombres
                })
                otros_mujeres = pd.DataFrame({
                    "Sexo": ["Mujer"] * len(detalle_motivos_otros_mujeres),
                    "Detalle_Otros": detalle_motivos_otros_mujeres
                })
                df_otros = pd.concat([otros_hombres, otros_mujeres], ignore_index=True)

                # Estadísticas generales
                resumen = {
                    "Total consultas": [hombres + mujeres],
                    "Total hombres": [hombres],
                    "Total mujeres": [mujeres],
                    "Hombres < 60 años": [sum(1 for e, _ in datos_hombres if e != -1 and e < 60)],
                    "Hombres >= 60 años": [sum(1 for e, _ in datos_hombres if e != -1 and e >= 60)],
                    "Hombres missing": [sum(1 for e, _ in datos_hombres if e == -1)],
                    "Mujeres < 60 años": [sum(1 for e, _ in datos_mujeres if e != -1 and e < 60)],
                    "Mujeres >= 60 años": [sum(1 for e, _ in datos_mujeres if e != -1 and e >= 60)],
                    "Mujeres missing": [sum(1 for e, _ in datos_mujeres if e == -1)],
                }
                df_resumen = pd.DataFrame(resumen)

                # Pestaña distribución por edad, motivo y sexo
                def clasifica_edad(edad):
                    if edad == -1:
                        return ">=80 o desconocida"
                    elif edad < 60:
                        return "<60"
                    else:
                        return ">=60"

                df['Grupo_Edad'] = df['Edad'].apply(clasifica_edad)
                dist = df.groupby(['Sexo', 'Grupo_Edad', 'Motivo']).size().reset_index(name='Cuenta')
                dist = dist.sort_values(['Sexo', 'Grupo_Edad', 'Motivo'])

                # Guardar en Excel con varias hojas
                nombre_archivo = f"analisis_colon_{anyo if anyo else 'todos'}.xlsx"
                with pd.ExcelWriter(nombre_archivo) as writer:
                    df.to_excel(writer, sheet_name="Datos", index=False)
                    df_otros.to_excel(writer, sheet_name="Otros motivos", index=False)
                    df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
                    dist.to_excel(writer, sheet_name="Distribución", index=False)

                print(f"\nDatos exportados correctamente a '{nombre_archivo}'.")

# Opcional: Quitar el PST después de listar (descomentar si se desea)
# outlook.RemoveStore(root_folder)

