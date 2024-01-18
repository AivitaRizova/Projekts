# Automātiska Excel datu apstrāde un salīdzināšana
### uzdevums:
 - Izejvielu materiālu ienākumu automātiska aprēķināšana un salīdzināšana.

## Par projektu
 - Izstrādātajā Phyton programmā automātiski tiek apstrādāti un salīdzināti dati divos Excel failos. Tajos tiek aprēķināta summa rindām un kolonnām.
 - Rezultāti tiek saglabāti tajos pašos Excel failos.

## Izmantošanas metodes:
### 1. Instalēt nepieciešamās bibliotēkas:
- Pirms programmas izmantošanas ir jāinstalē nepieciešamā bibliotēka. To var izdarīt, izpildot komandas terminālī:
 > pip install openpyxl
### 2. Iegūt Excel failus:
- Pirms skripta izpildes pārliecinieties, ka jums ir divi Excel faili, kurus vēlaties apstrādāt un salīdzināt. Failu nosaukumi jāatbilst tiem, kas minēti jūsu skriptā ("izejvielas_materiali_T.xlsx" un "izejvielas_materiali_K.xlsx").
### 3. Izpildīt Phyton programmu:
- Atveriet termināli vai komandrindu un pārvietojieties uz to direktoriju, kurā atrodas jūsu Python skripts. Izpildiet skriptu.
### 4. Pārbaudiet rezultātus:
- Pēc skripta izpildes terminālī tiks izvadīti rezultāti, kas norāda, kurā no diviem failiem ir lielāka kopsumma. Tāpat rezultāti tiks saglabāti pašos Excel failos, un jūs varat tos pārbaudīt, atverot failus ar Excel lietojumprogrammu.
### 5.Pielāgojiet programmu(pēc nepieciešamības):
- Ja vēlaties veikt papildu aprēķinus vai salīdzinājumus, varat rediģēt skripta kodu un pielāgot to saviem vajadzības.
## Imantotās bibliotēkas un to izmantošana:
 - **openpyxl** -šī bibliotēka tiek izmantota darbam ar Excel failiem. **'load_workbook'** funkcija ļauj ielādēt esošos Excel dokumentus, **'Workbook'** funkcija ļauj izveidot jaunu dokumentu un **'get_column_letter'** tiek izmantota, lai iegūtu burtu, kas atbilst kolonnas numuram.

## Projekta posmi:
### 1.Datu nolasīšana:
  - Izmantojot **'load_workbook'**, tiek nolasīti dati no diviem Excel failiem ('izejvielas_materiali_T.xlsx' un 'izejvielas_materiali_K.xlsx').
  - Dati tiek nolasīti no Excel formulu rezultātiem, nevis no pašām formulām. To nodrošina **'data_only=True'**.

### 2. Datu apstrāde:
  - Abiem failiem tiek aprēķinātas summas rindām un kolonnām un rezultāti tiek ierakstīti atbilstošajās Excel šūnās.
  - Izmantojot **'for'** ciklus, tiek apstaigātas rindas un kolonnas, aprēķinot summas ar **'sum'** funkciju un ierakstot rezultātus atpakaļ Excel dokumentos.
  - Ar **'round'** noapaļo decimālos skaitļus līdz noteiktam skaitlim pēc punkta(šajā gadījuma 2).

### 3. Datu saglabāšana un aizveršana:
  - Lai saglabātu rezultējošos Excel failus, izmanto **'wb_t.save('izejvielas_materiali_T.xlsx')'** un **'wb_k.save('izejvielas_materiali_K.xlsx')'**.
 - Ar **'wb_t.close()'** un **'wb_k.close()'** tiek aizverti atvērtie Excel faili.

### 4. Datu salīdzināšana:
 - **'value_t'** un **'value_k'** iegūst summas no abiem failiem, kur ir attiecīgā kopsumma kolonna _'gadā'_ ('O15' un 'O16').
 - Izmantojot **'if-elif-else'** konstrukciju, tiek salīdzinātas summas, kuras ieguva no **'value_t'** un **'value_k'**, un terminalī izvadīts, kurā failā ir lielāka, mazāka, vai vienāda kopsumma.

