(function() {
    var params = new URLSearchParams(window.location.search);
    var imageUrl = params.get('image');
    var imgEl = document.getElementById('locationImage');
    var noImgEl = document.getElementById('noImage');
    if (imageUrl) {
        imgEl.src = imageUrl;
        imgEl.style.display = 'inline-block';
        imgEl.onerror = function() {
            imgEl.style.display = 'none';
            noImgEl.textContent = 'Image could not be loaded.';
        };
        noImgEl.style.display = 'none';
    }
    var fields = [
        { id: 'pcSection', param: 'pcSection' },
        { id: 'article', param: 'article' },
        { id: 'description', param: 'description' },
        { id: 'oldProperty', param: 'oldProperty' },
        { id: 'unitMeas', param: 'unitMeas' },
        { id: 'unitValue', param: 'unitValue' },
        { id: 'qty', param: 'qty' },
        { id: 'location', param: 'location' },
        { id: 'condition', param: 'condition' },
        { id: 'remarks', param: 'remarks' },
        { id: 'user', param: 'user' }
    ];
    fields.forEach(function(f) {
        var el = document.getElementById(f.id);
        var val = params.get(f.param);
        if (el) el.textContent = (val != null && val !== '') ? decodeURIComponent(val) : '—';
    });
})();
