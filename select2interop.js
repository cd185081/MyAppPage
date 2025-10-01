window.select2Interop = {
    init: function (id) {
        $('#' + id).select2({
            theme: 'bootstrap-5',
            placeholder: "Select options",
            closeOnSelect: false
            //tags: true,
            //tokenSeparators: [',', ' ']
        });
    },
    getValues: function (id) {
        return $('#' + id).val();
    }
};
