-- Crear base de datos
CREATE DATABASE HubLectura;
GO

-- Usar la base de datos
USE HubLectura;
GO

-- Crear tabla de usuarios
CREATE TABLE Usuarios (
    Id INT PRIMARY KEY IDENTITY(1,1),
    Nombre NVARCHAR(100),
    Preferencias NVARCHAR(200)
);
GO

-- Crear tabla de libros
CREATE TABLE Libros (
    Id INT PRIMARY KEY IDENTITY(1,1),
    Titulo NVARCHAR(100),
    Autor NVARCHAR(100),
    Genero NVARCHAR(50),
    Calificacion INT,
    Estado NVARCHAR(30), -- "Leído", "Quiero leer", etc.
    UsuarioId INT FOREIGN KEY REFERENCES Usuarios(Id)
);
GO


