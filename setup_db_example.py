import sqlite3
import pandas as pd
import os

def create_example_db():
    print("Creando base de datos de ejemplo...")
    
    # Crear carpeta si no existe
    os.makedirs(os.path.dirname(os.path.abspath(__file__)), exist_ok=True)
    
    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ejemplo_usuarios.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Crear tabla
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT,
        apellido TEXT,
        identificacion TEXT,
        direccion TEXT,
        telefono TEXT,
        correo TEXT,
        fecha TEXT
    )
    ''')
    
    # Limpiar tabla por si existe
    cursor.execute('DELETE FROM usuarios')
    
    # Insertar datos de ejemplo
    usuarios = [
        ('Juan', 'Perez', '12345', 'Calle 123', '555-1234', 'juan@example.com', '2023-10-01'),
        ('Maria', 'Gomez', '67890', 'Avenida 456', '555-5678', 'maria@example.com', '2023-10-02'),
        ('Carlos', 'Lopez', '11223', 'Carrera 789', '555-9012', 'carlos@example.com', '2023-10-03')
    ]
    
    cursor.executemany('''
    INSERT INTO usuarios (nombre, apellido, identificacion, direccion, telefono, correo, fecha)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', usuarios)
    
    conn.commit()
    
    # Demostración de consulta
    print("Exportando tabla a CSV para uso en la interfaz...")
    df = pd.read_sql("SELECT * FROM usuarios", conn)
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'usuarios_data.csv')
    df.to_csv(csv_path, index=False)
    
    conn.close()
    
    print(f"Base de datos creada en: {db_path}")
    print(f"Archivo CSV exportado en: {csv_path}")

if __name__ == "__main__":
    create_example_db()
