$('#startCopying').on('click', function(){
    startLoader('#step_1');
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            renderResponse(data, '#step_1');
        }
        stopLoader('#step_1')
    }
    $.post(url, response);
});

$('#startImport').on('click', function(){
    startLoader('#step_3');
    var url = $(this).data('url');
    function response(data) {
        if(data) {
            renderResponse(data, '#step_3');
        }
        stopLoader('#step_3')
    }
    $.post(url, response);
});

$('#uploadFileForm').on('click', function(event) {
    event.preventDefault();
    $('#step_2').find('.resp').remove();
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
        $('#step_2').append('<div class="resp">' + msg + '</div>');
    });
});

function renderResponse(data, node) {
    $(node).append('<div class="resp">' + data + '</div>');
}

function startLoader(node){
    $(node + ' div.resp').remove();
    $(node + ' div.step-body').css({'visibility':'hidden'});
    $(node).prepend('<div class="loading"></div>');
}

function stopLoader(node){
    $('.loading').remove();
    $(node + ' div.step-body').css({'visibility':'visible'});
}