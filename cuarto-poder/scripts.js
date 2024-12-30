document.addEventListener('DOMContentLoaded', () => {
    const sections = ['Cienaga', 'SantaMartaYMagdalena', 'Mundo', 'Opinion', 'Educacion', 'Sociedad', 'Deporte'];
    sections.forEach(section => loadArticles(section));
});

async function loadArticles(section) {
    const response = await fetch(`/articles/${section}`);
    const articles = await response.json();
    const articlesContainer = document.getElementById(`${section}-articles`);

    articlesContainer.innerHTML = articles.map(article => `
        <article>
            <img src="${article.image}" alt="Imagen de noticia">
            <h3>${article.title}</h3>
            <p>${article.content}</p>   
            <a href="article.html?section=${section}&title=${encodeURIComponent(article.title)}" target="_blank">Leer más</a>
        </article>
    `).join('');
}

// Menú móvil
document.getElementById('mobile-menu').addEventListener('click', () => {
    const navLinks = document.getElementById('nav-links');
    navLinks.classList.toggle('active');
});
