# -*- coding: utf-8 -*-
"""Dataframe e Serie di pandas-Mediolanum.ipynb

> (c) 2026 Antonio Michele Piemontese

Il package *pandas* di Python serve a gestire i **dataframe** e le **serie**.
"""
import display
import pandas as pd     # importazione in memoria del package pandas (già installato di default in Google Colab);
                        # non confondere la installazione di un package con la import: la prima installa SU DISCO e si fa una volta sola (alcuni package
                        # sono già pre-installati); la seconda carica il package IN MEMORIA e dunque va fatta ad ogni nuova esecuzione del notebook

df = pd.read_csv(filepath_or_buffer='Credit_ISLR.csv', sep=",")   # questa riga dà errore --> occorre fare l'upload del file csv nella session storage

"""Il file ***Credit*** è un famoso file bancario che contiene 400 clienti di carte di credito descritti lungo una decina di attributi. E' un file simulato negli anni '90 ed è molto usato per imparare python e pandas."""

df.head()               # le prime 5 righe

df.tail()               # le ultime 5 righe

"""---
pandas fornisce un insieme di funzioni chiamate **read_XXXX** per leggere file di molti formati. Vedi in proposito [questa chat](https://chatgpt.com/share/68e6379c-a2f4-8012-b0cd-615ff6a2ac74).

---

**Rimuoviamo** dal dataframe 'df' le colonne inutili:
- *Unnamed: 0*: inserita dalla funzione 'pd.read_csv'
- *ID*: un identificativo numerico del cliente presente nel file CSV. Numera da 1 anzichè da 0, come deve essere in Python; inoltre è ridondante rispetto all'indice che pandas crea in automatico (la prima colonna a sx in grassetto).
"""

df.drop(columns=['Unnamed: 0','ID'],inplace=True)  # in-place = True rende l'operazione di drop PERSISTENTE (in memoria)

display(df)       # display è una FUNZIONE, non un metodo

"""La funzione 'display' visualizza le righe e colonne di un dataframe in formato **HTML / CSS** (a differenza della funzione 'print').

Come si vede, le due colonne "droppate" non sono più presenti nel dataframe.

Per vedere **tutte** le righe di un dataframe dobbiamo **settare alcune impostazioni** di pandas (**solo per questo output**):
"""

with pd.option_context("display.max_rows", None, "display.max_columns", None):
    display(df)

"""La **classe** di un oggetto"""

type(df)              # la classe dell'oggetto --> è un dataframe perchè è stato creato con la funzione 'pd.read_csv' di pandas

"""I datatype elementari (delle singole colonne):"""

df.dtypes

"""Le dimensioni del dataframe con le relative cardinalità:"""

df.shape

"""L'occupazione in byte del dataframe:"""

df.size

"""**Cosa è un dataframe?**<br>
E' una TABELLA di righe e colonne **in memoria**. Non è la classica tabella SQL (su disco).

Vediamonne le **dimensioni** (con il relativo **size**):
"""

df.shape

"""Il dataframe 'df' ha 2 dimemnsioni di size 400 e 11.

L'elenco dei data-type pandas **elementari** del 'df' (cioè delle singole colonne) è fornito dal metodo 'dtypes'.
"""

df.dtypes

"""Quanto spazio in memoria (in byte) occupa il dataframe?"""

df.size

"""Il dataframe contiene valori mancanti? (detti **missing values** oppure **not available**). Si fa con il metodo 'isna' che funziona in questo modo:
- lavora in modo parrallelo su righe e colonne
- esegue il test sulla **singola cella**
- e quindi restituisce una matrice di booleani (True oppure False) con l'esito del test per ogni cella.
"""

df.isna()

"""Occorre individuare le celle con **True** (i MV). Per dataset medio-grandi questa operazione è di fatto impossibile.

Per fortuna, il valore booleano True è memorizzato nei PC come 1, ed il valore False come 0.
"""

df.isna().sum()

"""Ora un pò di **analisi** di questo dataset:"""

df.info()          # fornisce una SINTESI di informazioni sul dataframe

"""`float64` è un numero a virgola  mobile in doppia precisione (cioè con i decimali e allocato su 64 bit).<br>
`int64` è un numero intero allocato su 64 bit<br>
`object` è una stringa alfanumerica

Possiamo ottenere alcune di queste info separatamente, con comandi differenti:
"""

display(df.columns.tolist())
display(df.dtypes)
print('\n','numero di NA complessivo: ',df.isna().sum().sum().item(),'\n')
print('byte del size: ', df.size)

df.describe()   # calcola le statistiche di base per le SOLE colonne NUMERICHE (int, float, ecc)

round(df.describe(),2)

df['Income'].median()    # mediana molto più bassa della media; come mai? indaghiamo...

df['Income'].hist(bins=20)

"""---

**Nota sulla mediana**

Ad ottobre 2025 si è discusso molto su questo post Instagram di **StartingFinance**. Vedi [questa *reaction*](https://www.tiktok.com/@brainlink_project/video/7564921849491033366?is_from_webapp=1&sender_device=pc).<br>
![](stipendi_italiani.png)

> Commento di *StartingFinance*:<br>
> Lombardia e Trentino-Alto Adige si confermano le regioni con gli stipendi più alti d’Italia. Nel 2024, la RAL media è di circa €33.600, che corrisponde a €1.960 netti al mese per 13 mensilità. Seguono da vicino Lazio (€33.200) e Liguria (€33.100), mentre le regioni del Sud si fermano tra i €27.000 e i €29.000 lordi annui — come Basilicata, Calabria e Sicilia.<br>
> Questi dati, elaborati dall’Osservatorio JobPricing nel report Salary Outlook 2025, rappresentano valori medi. Significa che esistono differenze significative tra settori, ruoli e livelli di esperienza, ma servono a tracciare un quadro generale della distribuzione del reddito.<br>
> Le RAL regionali permettono infatti di stimare il potere d’acquisto e la competitività dei territori, ma anche di leggere il divario Nord-Sud: oltre €6.000 lordi all’anno separano le prime regioni in classifica dalle ultime.<br>
> Un indicatore imperfetto, certo — ma indispensabile per capire dove il lavoro “paga” di più.

Il metodo 'describe', come molti altri metodi e funzioni di pandas e python, utilizza il famoso **funzionamento VETTORIALE**, nel quale le colonne sono elaborate in modo **parallelo ed indipendente**.

In linguaggi tradizionali, anche ad oggetti come Java, C#, ecc, il metodo 'describe' - se applicato a tutto il dataframe - sarebbe implementato con un ciclo 'for' (una iterazione per colonna).

Il funzionamento vettoriale permette inoltre di scrivere **codice più compatto e leggibile**
"""

print(df.shape)
print(df['Income'].shape)
print(type(df['Income']))

print(df[['Income']].shape)
print(type(df[['Income']]))

"""pandas ha due tipi di data-type aggregati: il dataframe (righe e colonne) e **la serie**, composta da un vettore (serie di elementi).

Queste 2 classi hanno metodi differenti! (a parte alcuni: head, tail, shape, describe, ecc)
"""

df[['Income']].boxplot()

"""I baffi esprimono l'[IQR](https://it.wikipedia.org/wiki/Scarto_interquartile) (Inter Quartile Range), cioè la differenza tra Q3 (75%) e Q1 (25%), moltiplicato per 1.5. E' una convenzione che permette di inviduare eventuali outlier, cioè valori "anomali".

Calcoliamo la fondamentale matrice di correlazione
"""

df.corr()    # --> dà errore perchè le correlazioni sono calcolabili solo tra coppie di variabili numeriche

df.select_dtypes('number')

df.select_dtypes('object')

round(df.select_dtypes('number').corr(),2)

"""La matrice di correlazione è quadrata e simmetrica. Sulla diagonale principale ci sono tutti 1 -> la correlazione infatti è un coefficiente tra 0 e 1 (se positiva) e tra 0 e -1 (se negativa).

Vogliamo estrarre solo la matrice traingolare superiore.
"""

import numpy as np    #  un altro package di python fondamentale;
                      #  serve per molti calcoli numerici;
                      #  per convenzione è importato come 'np'

correlation_matrix = df.select_dtypes('number').corr()
upper_triangular_matrix = np.triu(correlation_matrix)
display(upper_triangular_matrix)   # numpy lavora internamente con le ARRAY (numeriche) e non i dataframe
                                   # quindi restituisce una array (una matrice numerica) SENZApiùi nomi delle colonne
mat_corr_upper = pd.DataFrame(upper_triangular_matrix, columns=df.select_dtypes('number').columns)
                                   # 'pd.DataFrame' è una FUNZIONE con la lista degli argomenti tra ();
                                   # invece 'select_dtypes' e 'columns' sono 2 METODI della classe DataFrame e si scrivono A VALLE dell'oggetto
display(round(mat_corr_upper,2))

"""La seguente cella è stata generata dall'**assistente AI**: avevamo chiesto come NON visualizzare la matrice triangolare inferiore (neanche gli 0)"""

# Hide the lower triangular part (including the diagonal)
def hide_lower_triangular(val):
    return 'visibility: hidden' if val == 0 else ''

# Apply the styling to the upper triangular matrix
styled_matrix = mat_corr_upper.style.map(hide_lower_triangular)
display(styled_matrix)

"""L'indice di correlazione tra Rating e Limit è altissimo (0.9968). Vuol dire che c'è una fortissima **associazione** tra le due variabili, non necessariamente un rapporto causa effetto. Un esempio migliore è la correlazione tra Balance e Limit (0.86): il limite è deciso dalla banca unicamente sulla base del Balance? Probabilmente, no.

Se ci fosse un rapporto causa-effetto tra due variabili correlate (ad esempio Balance e Limit), le troveremmo correlate in OGNI campione di questa popolazione di questa banca.

In QUESTO campione abbiamo rilevato una semplice ASSOCIAZIONE, che potrebbe benissimo non ripetersi in un altro campione.

In italiano e in inglese: correlazione vuol dire rapporto "causa-effetto". L'uomo della srrada intende così la correlazione.

**Rename** di un nome colonna
"""

df.rename(columns={'Balance': 'Balance_card'}, inplace=True)
df

"""**Campionamento casuale**:"""

df.sample(n=10)   # utile per esaminare il dataset

df.sample(n=10, random_state=1000) # con seme, per garantire la "riproducibilità dei risultati"

"""***Shuffling*** = rimescolamento (NON per serie temporali).<br>
E' buona norma RIMESCOLARE un dataframe prima di applicare ad esso un modello di Machine Learning, poichè il dataframe potrebbe essere ORDINATO rispetto a qualche colonna, e cioò può essere un problema per alcuni modelli.<br>
L'ispezione manuale (visiva) degli eventuali ordinamenti è possibile solo se il dataframe ha poche colonne.

"""

n = df.shape[0]                           # il numero di righe del dataframe
df_sample = df.sample(n,random_state=1)   # il trucco è qua, il campionamento è fatto su TUTTE le righe,
                                          # SENZA RI-IMMISSIONE del cliente estratto nell'urna (il dataframe),
                                          #  e quindi costituisce uno shuffling
display(df_sample)

"""Tabelle di contingenza (**frequenze**)"""

# df['Gender'].value_counts
df[['Gender']].value_counts()

"""Vediamo ora la **cross-reference**, cioè la tabella degli incroci tra due colonne (anche detta **tabella di conmtingenza a 2 fattori**.)"""

# questo NON è un METODO ma è una FUNZIONE
# un metodo si può applicare solo agli oggetti (cioè variabili) di una CERTA CLASSE,
# il concetto di funzione è più ampio e meno restrittivo
# 'pd.crosstab' è una FUNZIONE
pd.crosstab(index=df['Gender'],columns=df['Married'])

"""Gli **indici**:"""

df.index         # quello CORRENTE, non necessariamente quello INIZIALE creato automaticamente
                 # da pandas (un progressivo da 0 in poi)
                 # l'indice serve a velocizzare gli accessi.

"""Possiamo modificare l'indice, ad esempio per la colonna 'Education'"""

df.set_index('Education',inplace=True)
df

df.index        # fornisce l'elenco dei valori (tutti e 400!) del NUOVO indice corrente (Education)

df.index.is_unique   # test booleano sulla unicità dei valori dell'indice corrente

df.index.unique()    # lista dei valori dell'indice unici

"""Una colonna indicizzata non è più subsettabile:"""

df['Education'].head()

"""Il bottone "Spiega errore" ha fornito la seguente risposta:<br>
> *L'errore KeyError: 'Education' nella cella corrente si è verificato perché la colonna 'Education' è stata impostata come indice del DataFrame utilizzando **df.set_index('Education', inplace=True)**. Una volta che una colonna diventa l'indice, non è più accessibile come una normale colonna utilizzando la sintassi standard df['nome_colonna'].*

> *Per accedere all'indice, puoi usare **df.index**. Per visualizzare i primi valori dell'indice, puoi usare **df.index.head()**.*

> *Se hai bisogno di accedere nuovamente a 'Education' come una normale colonna, puoi resettare l'indice usando **df.reset_index(inplace=True**).*
"""

df.index.head()      # suggerito da Gemini sopra --> un'allucinazione!

"""Uno dei vantaggi dell'argomento 'inplace=True' (disponibile per molti metodi pandas) è di **evitare la proliferazione inutile di dataframe**. Un'alternativa sarebbe infatti:
df2 = df.copy()
"""

df2 = df       # non crea una copia, crea solo un secondo puntatore al medesimo dataframe - può creare problemi, nel senso che, se modifico  'df' per qualche aspetto
               # risulta modificato anche 'df2'

df2 = df.copy  # crea una copia del primo dataframe, AUTONOMA --> consumo di memoria se i dataframe sono grandi

"""La proliferazione di dataframe inoltre confonde e rende più difficile la comprensione del codice.

**Quali sono i criteri per decidere la colonna (o le colonne) da indicizzare?**<br>
I criteri per decidere quale colonna del dataframe df ti conviene indicizzare dipendono principalmente da come intendi accedere ai dati. Ecco i punti chiave da considerare:
- **Accesso frequente e veloce per valore**: Se prevedi di cercare o recuperare righe molto spesso in base ai valori di una specifica colonna (ad esempio, cercare clienti per ID, o prodotti per codice), rendere quella colonna l'indice può velocizzare notevolmente queste operazioni. Gli indici in pandas funzionano in modo simile agli indici nei database, creando una struttura dati ottimizzata per le ricerche.
- **Unicità dei valori**: Se la colonna contiene valori univoci per ogni riga (come un ID cliente o un codice fiscale), è un ottimo candidato per diventare l'indice. Un indice univoco garantisce che ogni riga sia identificata in modo non ambiguo.
- **Ordinamento dei dati**: Se hai bisogno di accedere frequentemente a intervalli di righe basati sui valori di una colonna (ad esempio, tutte le transazioni in un determinato intervallo di date se la colonna è una data), impostare quella colonna come indice e ordinarla può migliorare le prestazioni.
- **Join e merge** (RARO): Quando esegui operazioni di join o merge tra dataframe, se i dataframe sono indicizzati sulla colonna che usi per il join, l'operazione sarà generalmente più veloce.
- **Significato dei dati**: A volte, una colonna ha un significato intrinseco che la rende un buon candidato per l'indice, anche se non è strettamente necessario per le prestazioni. Ad esempio, in un dataframe di serie temporali, la colonna della data/ora è quasi sempre l'indice naturale.

**Quando NON conviene indicizzare?**
- Se la colonna ha **molti valori duplicati** e non hai bisogno di accedere rapidamente a gruppi specifici basati su quei valori.
- Se modifichi frequentemente l'indice del dataframe, poiché la creazione e la gestione dell'indice hanno **un costo computazionale**.
- Se **l'ordine delle righe ORIGINALE è importante** e non vuoi che venga modificato dall'ordinamento implicito che può derivare dall'impostazione di un indice.

Nel caso del tuo dataframe 'df', colonne come 'ID' (se fosse ancora presente e univoca) o forse 'Education' (se volessi raggruppare o accedere spesso per livello di istruzione) potrebbero essere considerate come indici, **a seconda dell'analisi che intendi svolgere**.

> Tuttavia, per molte analisi esplorative di base, l'indice numerico di default è perfettamente adeguato.

Possiamo resettare l'indice a quello originale:
"""

df.reset_index(inplace=True)  # come sempre, l'argomento 'inplace=True' NON visualizza nulla ma modifica l'oggetto in memoria

df.head()

"""Cambiamo argomento, vediamo i metodi *style*:

Il metodo `style`, che ha molti **sottometodi**; vediamo i principali:
"""

df.select_dtypes('number').head().style.map(lambda v: 'color: red' if v < 20 else '')
                                                 # questa funzione lambda evidenzia in ROSSO i valori (di tutte le colonne NUMERICHE) < 20)
                                                 # notate l'applicazione IN CASCATA di 3 metodi

"""Una applicazione più frequente del metodo 'style' è **sulla singola colonna con cutoff specifico per colonna**
:
"""

df[['Income']].head().style.map(lambda v: 'color: red' if v > 100 else '')

"""NB. df[['Income']] è un dataframe. Invece,df['Income'] è una serie, che non dispone del metodo style --> errore!!"""

df['Income'].head().style.map(lambda v: 'color: red' if v > 100 else '')   # --> Series' object has no attribute 'style'

"""Altri tipi di uso di style:"""

display(df.select_dtypes('number').head(10).style.bar(subset=['Income'], color='lightblue'))
df.select_dtypes('number').head(10).style.bar(subset=['Income', 'Balance_card'], color='#d65f5f')

df.head().style.background_gradient(cmap='coolwarm')  # lo schema di colori qui scelto visualizza, per OGNI colonna NUMERICA del dataframe, in gradazioni di rosso i valori
                                                      # più alti e in gradazioni di blu quelli più bassi (non in assoluto, ma relativamente ad ogni colonna)

df.select_dtypes('number').style.highlight_max(axis=0,color='red')  # identifica (qui in rosso) il valore MAX di ogni colonna

"""Con il metodo `style` possiamo ora **visualizzare meglio la matrice di correlazione**:"""

df.select_dtypes('number').corr().style.background_gradient(cmap='coolwarm')

"""**Nota sul metodo 'style'**:<br>
La valutazione 'alto' o 'basso' è sempre **relativa** agli altri valori di QUELLA colonna singola, valutata **indipendentemente** dalle altre.

Cambiamo argomento. Passiamo al **subsetting**, cioè l'estrazione da un dataframe o da una serie pandas di un **sottoinsieme di righe e/o colonne**.<br>
Abbiamo già visto numerosi esempi di **subsetting di colonna**.<br>
Vediamo ora il subsetting di **riga**:
"""

df[1:3]        # attenzione a 2 cose: a. python conta da 0, b. l'estremo superiore non è (n), ma (n+1)

df['Income'][1:3]        # subsetting di colonna e riga

"""> Il subsetting permette di **focalizzare** l'analisi su una **porzione** del dataframe.

**Trasformazione** di una variabile categorica (da numerica a fattore).<br>
La variabile 'gender' (come anche le variabili 'Married', 'Student', ecc) è stata caricata in pandas come 'object' (stringa). E' meglio trasformare le variabili 'gender', 'married', ecc in variabili CATEGORICHE.<br>
Nella Data Science ci sono infatti due tipi di variabili:
- numeriche
- categoriche

La differenza è che le prime possono assumere un numero INFINITO di valori, le secondo un numero FINITO.
"""

df['Gender'] = df['Gender'].astype('category')
df.dtypes

df.select_dtypes(['category'])

"""**Ordinamenti**:"""

df.sort_values(by = 'Income', ascending = True,inplace=False) # senza 'inplace=True', altrimenti diventa permanente;

df.head()

"""**Raggruppamenti**: solo sulle colonne <u>numeriche</u>, applicando ad ogni gruppo una <u>funzione statistica</u> (media, mediana, ecc)"""

df.select_dtypes('number').groupby(by='Age').mean()

"""Se ai gruppi si applica un metodo non statistico (ad es. 'head()') la group_by non viene eseguita (provare).

Se ai gruppi non viene applicata nessuna funzione, non sono visualizzati risultati (provare).

La **standardizzazione** dei dati NUMERICI è spesso usata nella Data Science. Serve a trasformare una o più colonne NUMERICHE del dataframe nel range -6 +6 con media 0. Cioè, in altre parole, serve a **ricondurre tutte le colonne ad una stessa scala**.

La seguente cella utilizza la famosa libreria [**scikit-learn**](https://scikit-learn.org/stable/) per il Machine Learning in Python e la funzione *scale* del modulo *pre-processing*:
"""

import numpy as np                                        # numpy, pandas, matplotlib, ecc sono package STANDALONE e quindi si importano a se

from sklearn import preprocessing                         # scikit-learn è molto GRANDE e contiene molti moduli - si importa solo il modulo 'pre-processing'
                                                          # che contiene MOLTE funzioni di pre-elaborazione


np.set_printoptions(suppress=False)                       # sopprime l'uso della notazione scientifica per piccoli numeri: # imposta la precisione dell'output della mantissa
                                                          # (se tutte le cifre decimali danno fastidio).

arr_std = preprocessing.scale(df.select_dtypes('number')) # la standardizzazione si può fare solo su colonne numeriche

print(type(arr_std))                                      # numpy ha preso un dataframe pandas in input e ha restituito una array numpy
print(arr_std.shape)                                      # il metodo 'shape' (le dimensioni dell'oggetto) è uno dei pochi disponibili sia in pandas che numpy

arr_std                                                   # --> l'elenco dei 400 clienti (con 7 colonne numeriche) STANDARDIZZATI

df.head()

"""Un dataset standardizzato NON è più comprensibile all'utente, ma è utile e comodo in molti casi. Vediamone uno:"""

df.boxplot()                    # boxplot comparato fianco a fianco di TUTTE le colonne NUMERICHE del dataframe
                                # un SOLO comando, senza ciclo di for, grazie al funzionamento VETTORIALE (cioè parallelo) di pandas

"""Poichè le scale delle varie colonne sono **molto differenti** la comparazione è impossibile. Un semplice modo per risolvere questo problema è visualizzare i dati standardizzati.

Siccome le array numpy non hanno il metodo 'boxplot' dobbiamo prima trasformare l'array numpy in un dataframe pandas.
"""

# creiamo un nuovo dataframe, i nomi di colonna li dobbiamo prendere dal dataframe numerico originario.
# le array numpy non hanno infatti i nomi colonna, nel senso che sono andati persi quando prima abbiamo chiamato
# la funzione 'scale'
df_comp = pd.DataFrame(arr_std, columns=df.select_dtypes('number').columns)
df_comp.boxplot()

"""I boxplot sono ora confrontabili perchè **sulla stessa scala**. Vediamo:
- quali colonne hanno  outlier
- quali colonne hanno una escursione di valori maggiore
- quali colonne sono distribuite in modo simmetrico oppure no
- ecc

---
Una array di numpy è un oggetto **multi-dimensionale** (un vettore, una matrice, un cubo --> un tensore), nel quale le celle hanno tutte lo stesso datatype.<br>
I dataframe pandas sono **eterogenei**, cioè ogni colonna può avere il suo data-type. Le array numpy sono **omogenee**, in genere numeriche.

---

Le **serie**
"""

s = pd.Series([10,2,23,4])
s

type(s)

s.shape

s.index

"""Le colonne di un dataframe sono 'serie'. In altri termini, un dataframe pandas è la somma di tante colonne:"""

type(df['Income'])

"""dataframe e serie hanno molti metodi differenti (alcuni in comuni)"""

display(s.shape)
display(df.shape)

s.boxplot  # ce l'ha solo il dataframe

"""# I Data Type
Le seguenti immagini descrivono, nell'ordine, i data type di **Python, Pandas e Numpy**, e li confrontano:
"""

Image('python_data_types.png') if IN_COLAB else display(Image(filename='python_data_types.png'))

Image('python_data_types2.png') if IN_COLAB else display(Image(filename='python_data_types2.png'))

Image('Series_vs_Dataframe.png') if IN_COLAB else display(Image(filename='Series_vs_Dataframe.png'))

Image('pandas_data_types.png') if IN_COLAB else display(Image(filename='pandas_data_types.png'))

Image('numpy_array.png') if IN_COLAB else display(Image(filename='numpy_array.png'))

"""La seguente immagine confronta un dataframe con una array."""

Image('df_vs_array.png') if IN_COLAB else display(Image(filename='df_vs_array.png'))

"""Segue nella sottostante un confronto dei **data type elementari (atomici)** di pandas, python e numpy.<br>
Sono esclusi da questo confronto i **data type strutturati**, cioè:
- `dataframe` e `serie` per *pandas*
- `array` per *numpy*
- `set`, `dict` e `sequence` per *python*
"""

Image('data_types_mapping.png') if IN_COLAB else display(Image(filename='data_types_mapping.png'))







