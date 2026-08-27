from typing import List, Optional
from pydantic import BaseModel, Field, field_validator

class Egresado(BaseModel):
    num_egresado: str = ""
    apellido_nombre: str = ""
    documento: str = ""
    estado: str = "Aprobado"  # Aprobado, Ausente, Desaprobado

class TituloData(BaseModel):
    # Slide 1 fields
    apellido_nombre: str = Field(default="", description="Apellido y nombre del egresado")
    
    @field_validator("apellido_nombre", mode="after")
    @classmethod
    def force_uppercase(cls, v: str) -> str:
        return v.upper() if v else ""

    documento: str = Field(default="", description="DNI / Documento formateado (ej. 35.140.353)")
    horas_cursada: str = Field(default="", description="Cantidad de horas de cursada (ej. 230)")
    titulo_linea1: str = Field(default="", description="Nombre del título (Línea 1)")
    titulo_linea2: str = Field(default="", description="Nombre del título (Línea 2)")
    resolucion: str = Field(default="", description="Resolución Nro (ej. RESOC-2022-2450-GDEBA-DGCYE)")
    emision_dia: str = Field(default="", description="Día de emisión (ej. 02)")
    emision_mes: str = Field(default="", description="Mes de emisión (ej. Septiembre)")
    emision_ano: str = Field(default="", description="Año de emisión (ej. 25)")
    
    # Slide 2 fields
    modulos: List[str] = Field(default_factory=list, description="Lista de hasta 10 módulos con su código")
    fecha_egreso: str = Field(default="", description="Fecha de egreso (ej. 14 de Julio de 2025)")
    numero_egresado: str = Field(default="", description="Número de egresado (ej. 354)")
    numero_cpf: str = Field(default="412", description="Número CPF")
    distrito_cpf: str = Field(default="Lomas de Zamora", description="Distrito CPF")
