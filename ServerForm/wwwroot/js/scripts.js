$(document).ready(function () {
    // Обработчик события для кнопки
    $('#myButton').click(function () {
        alert('Кнопка нажата!');
    });

    // Пример функции для загрузки данных через AJAX
    $('#loadData').click(function () {
        $.ajax({
            url: '/api/data', // URL вашего API
            method: 'GET',
            success: function (data) {
                $('#dataContainer').html(data); // Отображение данных в контейнере
            },
            error: function () {
                alert('Ошибка при загрузке данных.');
            }
        });
    });
});