<?php
session_start();

function login() {
    $username = $_POST['username'];
    $password = $_POST['password'];

    $stored_username = 'admin';
    $stored_password_hash = '$2a$10$YSJfPqhpM0xtrEq0VvtiOebRx6Q2DC.hsxoDO4nQjfw6F09PzZik2'; // Hash de 'password980119301'

    if ($username === $stored_username && password_verify($password, $stored_password_hash)) {
        $_SESSION['user'] = $username;
        echo json_encode(['token' => session_id()]);
    } else {
        http_response_code(401);
        echo json_encode(['message' => 'Credenciales inválidas']);
    }
}

function modify() {
    if (!isset($_SESSION['user'])) {
        http_response_code(401);
        echo json_encode(['message' => 'No autorizado']);
        exit();
    }

    $title = $_POST['title'];
    $oldTitle = $_POST['oldTitle'];
    $section = $_POST['section'];
    $content = $_POST['content'];
    $content2 = $_POST['content2'];
    $content3 = $_POST['content3'];
    $date = $_POST['date'];
    $uploadDir = __DIR__ . "/articles/$section/";

    if (!file_exists($uploadDir)) {
        mkdir($uploadDir, 0777, true);
    }

    $oldTitle_2 = $_POST['oldTitle_2'];
    $oldTitle_3 = $_POST['oldTitle_3'];
    $oldTitle_date = $_POST['oldTitle_date'];
    
    if ($oldTitle && $oldTitle !== $title) {
        @unlink("$uploadDir$oldTitle.txt");
        @unlink("$uploadDir$oldTitle.jpg");
        @unlink("$uploadDir$oldTitle_2.txt");
        @unlink("$uploadDir$oldTitle_2.jpg");
        @unlink("$uploadDir$oldTitle_3.txt");
        @unlink("$uploadDir$oldTitle_3.jpg");
        @unlink("$uploadDir$oldTitle_date.txt");
        @unlink("$uploadDir$oldTitle.mp4");
    }

    file_put_contents("$uploadDir$title.txt", $content);
    file_put_contents("$uploadDir{$title}_date.txt", $date ?: date('Y-m-d'));

    if ($content2) {
        file_put_contents("$uploadDir{$title}_2.txt", $content2);
    }
    if ($content3) {
        file_put_contents("$uploadDir{$title}_3.txt", $content3);
    }

    foreach (['image', 'image2', 'image3', 'video'] as $key) {
        if (!empty($_FILES[$key]['name'])) {
            move_uploaded_file($_FILES[$key]['tmp_name'], "$uploadDir$title" . ($key === 'image' ? '' : "_$key") . "." . pathinfo($_FILES[$key]['name'], PATHINFO_EXTENSION));
        }
    }

    echo json_encode(['message' => 'Contenido modificado']);
}

function deleteArticle() {
    if (!isset($_SESSION['user'])) {
        http_response_code(401);
        echo json_encode(['message' => 'No autorizado']);
        exit();
    }

    $title = $_POST['title'];
    $section = $_POST['section'];
    $uploadDir = __DIR__ . "/articles/$section/";

    @unlink("$uploadDir$title.txt");
    @unlink("$uploadDir$title.jpg");
    @unlink("$uploadDir{$title}_2.txt");
    @unlink("$uploadDir{$title}_2.jpg");
    @unlink("$uploadDir{$title}_3.txt");
    @unlink("$uploadDir{$title}_3.jpg");
    @unlink("$uploadDir{$title}_date.txt");
    @unlink("$uploadDir$title.mp4");

    echo json_encode(['message' => 'Artículo eliminado']);
}

function getArticles() {
    $section = $_GET['section'];
    $uploadDir = __DIR__ . "/articles/$section/";

    $articles = [];
    if (file_exists($uploadDir)) {
        foreach (glob("$uploadDir*.txt") as $file) {
            if (strpos($file, '_2.txt') === false && strpos($file, '_3.txt') === false && strpos($file, '_date.txt') === false) {
                $title = basename($file, '.txt');
                $content = file_get_contents($file);
                $date = file_exists("$uploadDir{$title}_date.txt") ? file_get_contents("$uploadDir{$title}_date.txt") : 'Fecha no disponible';
                $image = file_exists("$uploadDir$title.jpg") ? "/articles/$section/$title.jpg" : '';
                $content2 = file_exists("$uploadDir{$title}_2.txt") ? file_get_contents("$uploadDir{$title}_2.txt") : '';
                $image2 = file_exists("$uploadDir{$title}_2.jpg") ? "/articles/$section/{$title}_2.jpg" : '';
                $content3 = file_exists("$uploadDir{$title}_3.txt") ? file_get_contents("$uploadDir{$title}_3.txt") : '';
                $image3 = file_exists("$uploadDir{$title}_3.jpg") ? "/articles/$section/{$title}_3.jpg" : '';
                $video = file_exists("$uploadDir$title.mp4") ? "/articles/$section/$title.mp4" : '';

                $articles[] = [
                    'title' => $title,
                    'content' => $content,
                    'date' => $date,
                    'image' => $image,
                    'content2' => $content2,
                    'image2' => $image2,
                    'content3' => $content3,
                    'image3' => $image3,
                    'video' => $video,
                ];
            }
        }
    }

    echo json_encode($articles);
}

function getAdditionalContent() {
    $additionalContentDir = __DIR__ . '/additionalContent/';
    $contentPath = "$additionalContentDir/additionalContent.txt";
    $socialMediaContent1Path = "$additionalContentDir/socialMediaContent1.txt";
    $socialMediaContent2Path = "$additionalContentDir/socialMediaContent2.txt";
    $socialMediaImage1Path = "$additionalContentDir/socialMediaImage1.jpg";
    $socialMediaImage2Path = "$additionalContentDir/socialMediaImage2.jpg";

    $content = file_exists($contentPath) ? file_get_contents($contentPath) : '';
    $socialMediaContent1 = file_exists($socialMediaContent1Path) ? file_get_contents($socialMediaContent1Path) : '';
    $socialMediaContent2 = file_exists($socialMediaContent2Path) ? file_get_contents($socialMediaContent2Path) : '';
    $socialMediaImage1 = file_exists($socialMediaImage1Path) ? "/additionalContent/socialMediaImage1.jpg" : '';
    $socialMediaImage2 = file_exists($socialMediaImage2Path) ? "/additionalContent/socialMediaImage2.jpg" : '';

    echo json_encode([
        'content' => $content,
        'socialMediaContent1' => $socialMediaContent1,
        'socialMediaContent2' => $socialMediaContent2,
        'socialMediaImage1' => $socialMediaImage1,
        'socialMediaImage2' => $socialMediaImage2,
    ]);
}

function modifyAdditionalContent() {
    if (!isset($_SESSION['user'])) {
        http_response_code(401);
        echo json_encode(['message' => 'No autorizado']);
        exit();
    }

    $content = $_POST['content'];
    $socialMediaContent1 = $_POST['socialMediaContent1'];
    $socialMediaContent2 = $_POST['socialMediaContent2'];
    $additionalContentDir = __DIR__ . '/additionalContent/';

    if (!file_exists($additionalContentDir)) {
        mkdir($additionalContentDir, 0777, true);
    }

    file_put_contents("$additionalContentDir/additionalContent.txt", $content);
    file_put_contents("$additionalContentDir/socialMediaContent1.txt", $socialMediaContent1);
    file_put_contents("$additionalContentDir/socialMediaContent2.txt", $socialMediaContent2);

    foreach (['socialMediaImage1', 'socialMediaImage2'] as $key) {
        if (!empty($_FILES[$key]['name'])) {
            move_uploaded_file($_FILES[$key]['tmp_name'], "$additionalContentDir$key.jpg");
        }
    }

    echo json_encode(['message' => 'Contenido adicional guardado']);
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_GET['action'])) {
        switch ($_GET['action']) {
            case 'login':
                login();
                break;
            case 'modify':
                modify();
                break;
            case 'delete':
                deleteArticle();
                break;
            case 'modifyAdditionalContent':
                modifyAdditionalContent();
                break;
        }
    }
} else if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (isset($_GET['action'])) {
        switch ($_GET['action']) {
            case 'getArticles':
                getArticles();
                break;
            case 'getAdditionalContent':
                getAdditionalContent();
                break;
        }
    }
}
?>