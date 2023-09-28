create database DB_Alumnos

use DB_Alumnos

Create table Calificaciones(
Matrícula nvarchar (15) not null,
Nombre nvarchar (100),
Español int,
Matemáticas int,
Biología int,
Historia int,
Inglés int,
Estado varchar(10),
Promedio decimal
)
