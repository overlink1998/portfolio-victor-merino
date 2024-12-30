const bcrypt = require('bcryptjs');

const password = 'password980119301';
const saltRounds = 10;

bcrypt.hash(password, saltRounds, (err, hash) => {
    if (err) {
        console.error('Error al generar el hash:', err);
    } else {
        console.log('Nuevo hash:', hash);
    }
});
