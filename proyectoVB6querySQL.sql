-- Si desea crear una base de datos nueva para almacenar la informacion generada, quite el comentario de la siguiente linea
-- CREATE DATABASE DB_VisionBoard

-- Usar la base de datos de su preferencia
-- USE DB_VisionBoard;

-- Crear la tabla 'propositos'
CREATE TABLE propositos (
    id INT IDENTITY(1,1) PRIMARY KEY,  
    imagen VARCHAR(255),
    titulo TEXT,
    categoria VARCHAR(255),
    fecha_terminacion DATE,
    descripcion TEXT
);

-- Insertar los registros en la tabla 'propositos'
INSERT INTO propositos (imagen, titulo, categoria, fecha_terminacion, descripcion) VALUES
('salud.jpg', 'Mejorar la salud f�sica', 'Salud', '2025-12-31', 'Hacer ejercicio regularmente, comer saludablemente y dormir mejor para mejorar el bienestar f�sico.'),
('leer.jpg', 'Leer m�s libros', 'Desarrollo personal', '2025-12-31', 'Leer al menos un libro al mes para expandir conocimientos y mejorar habilidades cognitivas.'),
('ahorro.jpg', 'Ahorrar dinero', 'Finanzas personales', '2025-12-31', 'Ahorrar un porcentaje de los ingresos cada mes y crear un fondo de emergencia.'),
('viajar.jpg', 'Viajar m�s', 'Aventura', '2025-12-31', 'Explorar nuevos lugares y culturas, hacer al menos un viaje importante durante el a�o.'),
('relaciones.jpg', 'Mejorar las relaciones personales', 'Relaciones', '2025-12-31', 'Pasar m�s tiempo con la familia y amigos, y mantener contacto regular con seres queridos.');
