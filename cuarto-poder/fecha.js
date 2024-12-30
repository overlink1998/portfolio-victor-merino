const fs = require('fs');
const path = require('path');

const sections = ['Cienaga', 'SantaMartaYMagdalena', 'Mundo', 'Opinion', 'Educacion', 'Sociedad', 'Deporte'];

sections.forEach(section => {
    const sectionDir = path.join(__dirname, 'public', 'articles', section);
    if (fs.existsSync(sectionDir)) {
        const files = fs.readdirSync(sectionDir).filter(file => file.endsWith('_date.txt'));
        files.forEach(file => {
            const filePath = path.join(sectionDir, file);
            const dateContent = fs.readFileSync(filePath, 'utf8');
            const dateOnly = new Date(dateContent).toISOString().split('T')[0];
            fs.writeFileSync(filePath, dateOnly);
        });
    }
});

console.log('Fechas actualizadas');
