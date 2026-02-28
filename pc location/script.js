(function() {
    var params = new URLSearchParams(window.location.search);
    var imageUrl = params.get('image') || params.get('img');
    // Support image in hash (long URLs: Excel may truncate query; hash keeps it)
    var hash = window.location.hash.slice(1);
    if (!imageUrl && hash) {
        var hashParams = new URLSearchParams(hash);
        imageUrl = hashParams.get('image') || hashParams.get('img');
    }
    if (imageUrl) {
        imageUrl = imageUrl.trim();
        try {
            if (imageUrl.indexOf('%') !== -1) imageUrl = decodeURIComponent(imageUrl);
        } catch (e) {}
    }
    var imgEl = document.getElementById('locationImage');
    var noImgEl = document.getElementById('noImage');
    if (imgEl && noImgEl && imageUrl && (imageUrl.startsWith('http://') || imageUrl.startsWith('https://') || imageUrl.startsWith('data:image/'))) {
        imgEl.src = imageUrl;
        imgEl.style.display = 'inline-block';
        noImgEl.style.display = 'none';
        imgEl.onerror = function() {
            imgEl.style.display = 'none';
            noImgEl.style.display = 'block';
            noImgEl.innerHTML = '<span>No image</span>';
        };
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
