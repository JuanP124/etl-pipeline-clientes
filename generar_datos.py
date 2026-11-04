import random
from datetime import datetime, timedelta
import pandas as pd
from faker import Faker

fake = Faker("es_CO")  # Generador con nombres y datos en español (Colombia)

def generar_lote(n=5000, sucursal="A"):
    registros = []
    for i in range(1, n + 1):
        # Generar datos base
        id_cliente = random.randint(1000, 99999)
        nombre = fake.name().upper() if sucursal == "A" else fake.name() # Mezclar mayúsculas para probar limpieza
        email = fake.email()
        
        # Fechas aleatorias en el último año
        dias_atras = random.randint(0, 365)
        fecha = (datetime.now() - timedelta(days=dias_atras)).strftime("%Y-%m-%d")
        
        # Montos con formato variado
        monto = round(random.uniform(50000, 5000000), 2)
        if sucursal == "A":
            monto_str = f"{monto:,.2f}" # Formato con comas/puntos
        else:
            monto_str = str(monto)

        estado = random.choice(["Activo", "Pendiente pago", "Cancelado", "nan", "", None])

        if sucursal == "A":
            registros.append({
                "id_cliente": id_cliente,
                "nombre_cliente": nombre,
                "email": email,
                "fecha_registro": fecha,
                "precio_producto": monto_str,
                "stock_disponible": estado,
            })
        else:
            registros.append({
                "ID_CLIENTE": id_cliente,
                "Nombre": nombre,
                "Correo Electronico": email,
                "Fecha": fecha,
                "Monto Venta": monto_str,
                "Estado": estado,
            })
            
    return pd.DataFrame(registros)

if __name__ == "__main__":
    print("Generando datos masivos para Sucursal A...")
    df_a = generar_lote(5000, "A")
    df_a.to_csv("muestra_sucursal_A.csv", index=False, encoding="utf-8-sig")
    
    print("Generando datos masivos para Sucursal B...")
    df_b = generar_lote(5000, "B")
    df_b.to_excel("muestra_sucursal_B.xlsx", index=False, engine="openpyxl")
    
    print("¡Archivos masivos generados con éxito!")