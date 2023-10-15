#%%
import pandas as pd
from difflib import SequenceMatcher
# %%
df = pd.read_excel("C:/Users/sesa736791/Documents/Excel/Reading/Copy of Asistentes de Supply On.xlsx", sheet_name="Errores a mostrar")
df = df.fillna("")
#%%
df["Company"].head(100)
#%%
def sigCompany(nombresCompany, i):
    opcion = input("¿El nombre {} está correctamente escrito? Digita N para reescribir el nombre: ".format(nombresCompany[i]))
    if opcion == "N":
        nombreCorrecto = input("Escribe el nombre correcto: ")
    else:
        nombreCorrecto = nombresCompany[i]
    return nombreCorrecto
#%%
df['Weight'] = ""
df["In"] = ""
#%%
companyCol = df.iloc[:, 6]
nombreCorrecto = sigCompany(companyCol, 0)

# Regresa la dimensionalidad del objeto, es decir, cuántos elementos hay en cada dimensión
for i in range(companyCol.shape[0]):
    flag = False

    df['Weight'][i] = SequenceMatcher(None, companyCol[i].title(), nombreCorrecto.title()).ratio()

    for word in companyCol[i].title().split():
        # print(companyCol[i].title().split(), nombreCorrecto.title().split())
        if(word in nombreCorrecto.title().split()):
            flag = True
    
    df['In'][i] = flag
    # print(i, companyCol[i], nombreCorrecto, flag, df['Weight'][i])

    # Condicional donde si la distancia ponderada es menor a 0.50 y ninguna palabra se encuentra dentro del nombre estándar, pregunta por el sig company name
    if(df['Weight'][i] < 0.50 and not flag):
        nombreCorrecto = sigCompany(companyCol, i)
    
    df["Company"][i] = nombreCorrecto

    # Posible issue de que el nombre de la compañía sea similar al nombre de la siguiente company

#%%
import seaborn as sns
import matplotlib.pyplot as plt

plt.figure(figsize=(10, 6))  # Adjust the figure size if needed
sns.histplot(data=df, x="Weight", bins=16)
plt.title(f'Histogram of Weight')
plt.xlabel("Weight")
plt.ylabel('Frequency')
plt.show()
#%%
df[["Company","Weight", "In"]]
# %%
from datetime import datetime

df.to_excel(r"C:/Users/sesa736791/Documents/Excel/Created/Company Names Standarized {}.xlsx".format(datetime.now().strftime("%d-%m-%Y")), index=False)
# %%
# %%
