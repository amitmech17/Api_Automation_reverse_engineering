const crypto = require('crypto');

function username_encrypted(username) {
    const algorithm = 'aes-256-cbc';
    const key = 'GB9h2C9g0SN6YaOj3x5inT6KUEwoAzTG';
    const iv = 'YWzAoPjp26QGtJDt';

    const cipher = crypto.createCipheriv(algorithm, key, iv);

    const dataToEncrypt = username;

    var encryptedData = cipher.update(dataToEncrypt, 'utf8', 'hex');

    encryptedData = cipher.final('hex');

    console.log(encryptedData);
    //return encryptedData;
}

function password_hass(pass) {
    const password_hassed = crypto.pbkdf2Sync(pass,
        '865ab5ce-0d67-11e8-ba89-0ed5f89f718b', 10000, 64, 'sha512');
    console.log(password_hassed.toString('hex'));
    //return password_hassed;
}

// Process command-line arguments
const args = process.argv.slice(2); // Slice to remove first two elements (node executable and script path)

// Check if arguments are provided
if (args.length >= 2) {
    const number = args[0];
    const password = args[1];
    username_encrypted(number);
    password_hass(password);
} else {
    console.error('Usage: node your_script.js <command> <argument>');
    console.error('Commands:');
    console.error('  encrypt <username> - Encrypts the given username');
    console.error('  hash <password>   - Hashes the given password');
}
