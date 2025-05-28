# Anàlisi Formulari Web Còlon

## Introducció

Aquest projecte analitza un .pst per trovar tots els mails inclosos en una carpeta específica que inclouen totes les respostes que s'han anat generant al formulari web de Còlon https://icoprevencio.cat/colon/ al llarg dels anys. Amb això genera una distribució de les dades trovades.

Per executar aquest programari es presuposen uns coneixements mínims de la terminal de Windows.

## Passos previs

En aquest repositori públic no està inclòs el .pst ja que té informació sensible. L'usuari que executi el programa analisi.py s'haura de generar un .pst de la carpeta FORMULARI WEB de la bùstia compartida ICO Prevenció COLON. Això es fa amb l'exportació inclosa a l'Outlook 365 per escriptori.

Aquest fitxer s'haura de nomenar **colon.pst** i estar a l'arrel d'on està el programa **analisi.py**

## Requisitis

Python 3.13.3 https://www.python.org/ftp/python/3.13.3/python-3.13.3-amd64.exe

Instalació dels mòduls pywin32, pandas i openpyxl:

```bash
python -m pip install pywin32 pandas openpyxl
```

## Execució

Una vegada obtingut el fitxer .pst es farà doble click a **analisi.py** perque s'executi. Primer demanarà l'any:

```bash
Introduce el año (formato 4 dígitos, dejar vacío para todos):
```

Deixem buit si volem tots els anys o introduim un, per exemple 2025.

Després ens demanarà si volem veure el 'alres motius'

```bash
¿Quieres ver el detalle de los 'otros motivos'? (s/n):
```
Podem contestar **s** o **n**

Per acabar ens generarà un excel amb les dades seleccionades si contestem que sí a l'última pregunta

```bash
¿Quieres exportar los datos a excel (s/n)?
```