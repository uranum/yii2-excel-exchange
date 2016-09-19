$('#startCopying').on('click', function(){
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            renderResponse(data, '#step_1');
        }
    }
    $.post(url, response);
});
$('#startImport').on('click', function(){
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            renderResponse(data, '#step_3');
        }
    }
    $.post(url, response);
});

$('#uploadFileForm').on('click', function(event) {
    event.preventDefault();
    var url = $(this).data('url');
    var fd = new FormData();
    fd.append("ImportXls[file]", document.getElementById('uploadFile').files[0]);

    $.ajax({
        url: url,
        type: "POST",
        data: fd,
        processData: false,
        contentType: false
    }).done(function(msg) {
        renderResponse(msg, '#step_2');
    });
});

function renderResponse(data, elem) {
    $(elem).empty();
    $(elem).append(data);
}
