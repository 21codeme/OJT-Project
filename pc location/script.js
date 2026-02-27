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
        { id: 'building', param: 'building' },
        { id: 'room', param: 'room' },
        { id: 'location', param: 'location' },
        { id: 'units', param: 'units' },
        { id: 'condition', param: 'condition' },
        { id: 'remarks', param: 'remarks' },
        { id: 'updated', param: 'updated' }
    ];
    fields.forEach(function(f) {
        var el = document.getElementById(f.id);
        var val = params.get(f.param);
        if (el) el.textContent = (val != null && val !== '') ? decodeURIComponent(val) : '—';
    });
})();
