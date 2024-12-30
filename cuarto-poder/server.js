const express = require('express');
const bodyParser = require('body-parser');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Middleware
app.use(bodyParser.json());
app.use(fileUpload());
app.use(express.static(path.join(__dirname)));

// Reemplaza esto con el hash que generaste
const passwordHash = '$2a$10$YSJfPqhpM0xtrEq0VvtiOebRx6Q2DC.hsxoDO4nQjfw6F09PzZik2'; // Hash de 'password980119301'

// Base de datos de usuarios falsa (en una aplicación real, usa una base de datos)
const users = [{ username: 'admin', password: passwordHash }];

// Secreto para JWT
const jwtSecret = 'miClaveSecreta';

// Middleware de autenticación
const authenticate = (req, res, next) => {
    const token = req.headers['authorization'];
    if (token) {
        jwt.verify(token, jwtSecret, (err, decoded) => {
            if (err) {
                return res.status(401).json({ message: 'No autorizado' });
            } else {
                req.user = decoded;
                next();
            }
        });
    } else {
        return res.status(401).json({ message: 'No se proporcionó un token' });
    }
};

// Ruta de inicio de sesión
app.post('/login', (req, res) => {
    const { username, password } = req.body;
    const user = users.find(u => u.username === username);

    if (!user || !bcrypt.compareSync(password, user.password)) {
        return res.status(401).json({ message: 'Credenciales inválidas' });
    }

    const token = jwt.sign({ username: user.username }, jwtSecret, { expiresIn: '1h' });
    res.json({ token });
});

// Ruta protegida para agregar o actualizar contenido
app.post('/modify', authenticate, (req, res) => {
    const { content, title, section, oldTitle, content2, content3, date } = req.body;
    const image = req.files ? req.files.image : null;
    const image2 = req.files ? req.files.image2 : null;
    const image3 = req.files ? req.files.image3 : null;
    const video = req.files ? req.files.video : null;

    // Directorio donde se almacenan los artículos
    const sectionDir = path.join(__dirname, 'articles', section);
    if (!fs.existsSync(sectionDir)) {
        fs.mkdirSync(sectionDir, { recursive: true });
    }

    // Si se proporciona oldTitle, elimina el artículo antiguo
    if (oldTitle && oldTitle !== title) {
        const oldContentPath = path.join(sectionDir, `${oldTitle}.txt`);
        const oldImagePath = path.join(sectionDir, `${oldTitle}.jpg`);
        const oldContent2Path = path.join(sectionDir, `${oldTitle}_2.txt`);
        const oldImage2Path = path.join(sectionDir, `${oldTitle}_2.jpg`);
        const oldContent3Path = path.join(sectionDir, `${oldTitle}_3.txt`);
        const oldImage3Path = path.join(sectionDir, `${oldTitle}_3.jpg`);
        const oldDatePath = path.join(sectionDir, `${oldTitle}_date.txt`);
        const oldVideoPath = path.join(sectionDir, `${oldTitle}.mp4`);
        if (fs.existsSync(oldContentPath)) fs.unlinkSync(oldContentPath);
        if (fs.existsSync(oldImagePath)) fs.unlinkSync(oldImagePath);
        if (fs.existsSync(oldContent2Path)) fs.unlinkSync(oldContent2Path);
        if (fs.existsSync(oldImage2Path)) fs.unlinkSync(oldImage2Path);
        if (fs.existsSync(oldContent3Path)) fs.unlinkSync(oldContent3Path);
        if (fs.existsSync(oldImage3Path)) fs.unlinkSync(oldImage3Path);
        if (fs.existsSync(oldDatePath)) fs.unlinkSync(oldDatePath);
        if (fs.existsSync(oldVideoPath)) fs.unlinkSync(oldVideoPath);
    }

    // Guarda el contenido y la fecha
    const contentPath = path.join(sectionDir, `${title}.txt`);
    fs.writeFileSync(contentPath, content);

    const datePath = path.join(sectionDir, `${title}_date.txt`);
    fs.writeFileSync(datePath, date || new Date().toISOString().split('T')[0]);

    if (content2) {
        const content2Path = path.join(sectionDir, `${title}_2.txt`);
        fs.writeFileSync(content2Path, content2);
    }
    if (content3) {
        const content3Path = path.join(sectionDir, `${title}_3.txt`);
        fs.writeFileSync(content3Path, content3);
    }

    // Guarda las imágenes si se cargan
    if (image) {
        const imagePath = path.join(sectionDir, `${title}.jpg`);
        image.mv(imagePath, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga de la imagen', error: err });
            }
        });
    }
    if (image2) {
        const image2Path = path.join(sectionDir, `${title}_2.jpg`);
        image2.mv(image2Path, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga de la imagen 2', error: err });
            }
        });
    }
    if (image3) {
        const image3Path = path.join(sectionDir, `${title}_3.jpg`);
        image3.mv(image3Path, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga de la imagen 3', error: err });
            }
        });
    }

    // Guarda el video si se carga
    if (video) {
        const videoPath = path.join(sectionDir, `${title}.mp4`);
        video.mv(videoPath, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga del video', error: err });
            }
        });
    }

    res.json({ message: 'Contenido modificado', content });
});

// Ruta para eliminar un artículo
app.post('/delete', authenticate, (req, res) => {
    const { title, section } = req.body;

    // Directorio donde se almacenan los artículos
    const sectionDir = path.join(__dirname, 'articles', section);
    const contentPath = path.join(sectionDir, `${title}.txt`);
    const imagePath = path.join(sectionDir, `${title}.jpg`);
    const content2Path = path.join(sectionDir, `${title}_2.txt`);
    const image2Path = path.join(sectionDir, `${title}_2.jpg`);
    const content3Path = path.join(sectionDir, `${title}_3.txt`);
    const image3Path = path.join(sectionDir, `${title}_3.jpg`);
    const datePath = path.join(sectionDir, `${title}_date.txt`);
    const videoPath = path.join(sectionDir, `${title}.mp4`);

    if (fs.existsSync(contentPath)) fs.unlinkSync(contentPath);
    if (fs.existsSync(imagePath)) fs.unlinkSync(imagePath);
    if (fs.existsSync(content2Path)) fs.unlinkSync(content2Path);
    if (fs.existsSync(image2Path)) fs.unlinkSync(image2Path);
    if (fs.existsSync(content3Path)) fs.unlinkSync(content3Path);
    if (fs.existsSync(image3Path)) fs.unlinkSync(image3Path);
    if (fs.existsSync(datePath)) fs.unlinkSync(datePath);
    if (fs.existsSync(videoPath)) fs.unlinkSync(videoPath);

    res.json({ message: 'Artículo eliminado' });
});

// Ruta para servir la página principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Ruta para servir la página de administración
app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html'));
});

// Ruta para obtener todos los artículos
app.get('/articles/:section', (req, res) => {
    const section = req.params.section;
    const sectionDir = path.join(__dirname, 'articles', section);

    if (fs.existsSync(sectionDir)) {
        const articles = fs.readdirSync(sectionDir).filter(file => file.endsWith('.txt') && !file.includes('_2') && !file.includes('_3') && !file.includes('_date')).map(file => {
            const content = fs.readFileSync(path.join(sectionDir, file), 'utf8');
            const title = path.basename(file, '.txt');
            const datePath = path.join(sectionDir, `${title}_date.txt`);
            const date = fs.existsSync(datePath) ? fs.readFileSync(datePath, 'utf8') : 'Fecha no disponible';
            const image = fs.existsSync(path.join(sectionDir, `${title}.jpg`)) ? `/articles/${section}/${title}.jpg` : '';
            const content2Path = path.join(sectionDir, `${title}_2.txt`);
            const content2 = fs.existsSync(content2Path) ? fs.readFileSync(content2Path, 'utf8') : '';
            const image2 = fs.existsSync(path.join(sectionDir, `${title}_2.jpg`)) ? `/articles/${section}/${title}_2.jpg` : '';
            const content3Path = path.join(sectionDir, `${title}_3.txt`);
            const content3 = fs.existsSync(content3Path) ? fs.readFileSync(content3Path, 'utf8') : '';
            const image3 = fs.existsSync(path.join(sectionDir, `${title}_3.jpg`)) ? `/articles/${section}/${title}_3.jpg` : '';
            const video = fs.existsSync(path.join(sectionDir, `${title}.mp4`)) ? `/articles/${section}/${title}.mp4` : '';

            return { title, content, date, image, content2, image2, content3, image3, video };
        });
        res.json(articles);
    } else {
        res.json([]);
    }
});

// Ruta para servir la página de artículos
app.get('/articulo', (req, res) => {
    res.sendFile(path.join(__dirname, 'articulo.html'));
});

// Ruta para servir la página de noticias antiguas
app.get('/noticias-antiguas', (req, res) => {
    res.sendFile(path.join(__dirname, 'index2.html'));
});

// Ruta para modificar contenido adicional
app.post('/modifyAdditionalContent', authenticate, (req, res) => {
    const { content, socialMediaContent1, socialMediaContent2 } = req.body;
    const socialMediaImage1 = req.files ? req.files.socialMediaImage1 : null;
    const socialMediaImage2 = req.files ? req.files.socialMediaImage2 : null;

    // Directorio donde se almacenará el contenido adicional
    const additionalContentDir = path.join(__dirname, 'additionalContent');
    if (!fs.existsSync(additionalContentDir)) {
        fs.mkdirSync(additionalContentDir, { recursive: true });
    }

    
    // Guarda el contenido adicional
    const contentPath = path.join(additionalContentDir, 'additionalContent.txt');
    fs.writeFileSync(contentPath, content);

    const socialMediaContent1Path = path.join(additionalContentDir, 'socialMediaContent1.txt');
    fs.writeFileSync(socialMediaContent1Path, socialMediaContent1);

    const socialMediaContent2Path = path.join(additionalContentDir, 'socialMediaContent2.txt');
    fs.writeFileSync(socialMediaContent2Path, socialMediaContent2);

    // Guarda las imágenes si se cargan
    if (socialMediaImage1) {
        const socialMediaImage1Path = path.join(additionalContentDir, 'socialMediaImage1.jpg');
        socialMediaImage1.mv(socialMediaImage1Path, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga de la imagen de redes sociales 1', error: err });
            }
        });
    }
    if (socialMediaImage2) {
        const socialMediaImage2Path = path.join(additionalContentDir, 'socialMediaImage2.jpg');
        socialMediaImage2.mv(socialMediaImage2Path, err => {
            if (err) {
                return res.status(500).json({ message: 'Falló la carga de la imagen de redes sociales 2', error: err });
            }
        });
    }

    res.json({ message: 'Contenido adicional guardado' });
});

// Ruta para obtener el contenido adicional
app.get('/additionalContent', (req, res) => {
    const additionalContentDir = path.join(__dirname, 'additionalContent');
    const contentPath = path.join(additionalContentDir, 'additionalContent.txt');
    const socialMediaContent1Path = path.join(additionalContentDir, 'socialMediaContent1.txt');
    const socialMediaContent2Path = path.join(additionalContentDir, 'socialMediaContent2.txt');
    const socialMediaImage1Path = path.join(additionalContentDir, 'socialMediaImage1.jpg');
    const socialMediaImage2Path = path.join(additionalContentDir, 'socialMediaImage2.jpg');

    const content = fs.existsSync(contentPath) ? fs.readFileSync(contentPath, 'utf8') : '';
    const socialMediaContent1 = fs.existsSync(socialMediaContent1Path) ? fs.readFileSync(socialMediaContent1Path, 'utf8') : '';
    const socialMediaContent2 = fs.existsSync(socialMediaContent2Path) ? fs.readFileSync(socialMediaContent2Path, 'utf8') : '';
    const socialMediaImage1 = fs.existsSync(socialMediaImage1Path) ? `/additionalContent/socialMediaImage1.jpg` : '';
    const socialMediaImage2 = fs.existsSync(socialMediaImage2Path) ? `/additionalContent/socialMediaImage2.jpg` : '';

    res.json({ content, socialMediaContent1, socialMediaContent2, socialMediaImage1, socialMediaImage2 });
});

app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});
