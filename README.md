# App móvil de validación por consumo

Ejecutar:

```powershell
pip install -r requirements.txt
python -m streamlit run app.py
```

Correcciones principales:
- Kg alimento actual se interpreta como kg semanal.
- Kg diario actual = Kg semanal actual / Días alimentados.
- AO / Animales por alimento actual se calcula usando kg diario actual.
- BW sugerida usa NormalFrio para época Calor según validación contra el Excel mostrado.
