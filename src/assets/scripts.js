$('#startCopying').on('click', function(){
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            $('#step_1').empty();
            $('#step_1').append(data);
        }
    }
    $.post(url, response);
});

$('#startImport').on('click', function(){
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            $('#step_3').empty();
            $('#step_3').append(data);
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
        $('#step_2').empty();
        $('#step_2').append(msg);
    });
});