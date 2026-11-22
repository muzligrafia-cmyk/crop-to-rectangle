#target illustrator
/*
    CropToRectangle.jsx
    ------------------------------------------------------------
    Select one or more images with rectangles drawn over them,
    then run this script. Every rectangle crops the image under it:

      - the pixels outside the rectangle are permanently removed
      - each result is a new embedded image at 300 PPI
      - no clipping mask is used
      - the rectangles are deleted afterwards

    One image can have SEVERAL rectangles on it: each rectangle
    becomes its own cropped piece, and the original is deleted.
    (Overlapping rectangles are fine too.)

    Works on several images in one go.
    The document's colour mode (CMYK or RGB) is kept.
    ------------------------------------------------------------
*/

(function () {

    // ===== SETTINGS =====
    var RESOLUTION = 300;   // output resolution in PPI
    var TOL = 0.05;         // tolerance in points for "is this a straight rectangle?"
    // ====================

    if (app.documents.length === 0) {
        alert("Open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel;
    try { sel = doc.selection; } catch (e) { sel = null; }

    if (!sel || !sel.length) {
        alert("Select your images and the rectangles on top of them, then run the script again.");
        return;
    }

    var rects = [];
    var images = [];
    var notes = [];

    // ---------- helpers ----------

    function near(a, b) { return Math.abs(a - b) <= TOL; }

    // A closed, 4-corner, straight (unrotated), non-curved path
    function isAxisRect(p) {
        if (!p.closed || p.pathPoints.length !== 4) return false;
        var b = p.geometricBounds; // [left, top, right, bottom]
        for (var i = 0; i < 4; i++) {
            var pt = p.pathPoints[i];
            var a = pt.anchor;
            if (!(near(a[0], b[0]) || near(a[0], b[2]))) return false;
            if (!(near(a[1], b[1]) || near(a[1], b[3]))) return false;
            // no curve handles pulled out
            if (!near(pt.leftDirection[0], a[0]) || !near(pt.leftDirection[1], a[1])) return false;
            if (!near(pt.rightDirection[0], a[0]) || !near(pt.rightDirection[1], a[1])) return false;
        }
        return (b[2] - b[0]) > TOL && (b[1] - b[3]) > TOL;
    }

    // Walk the selection (and into groups) to find rectangles and images
    function collect(item) {
        try {
            if (item.locked || item.hidden) return;
        } catch (e) { return; }

        switch (item.typename) {
            case "GroupItem":
                for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
                break;
            case "PlacedItem":
            case "RasterItem":
                images.push(item);
                break;
            case "PathItem":
                if (item.clipping) return; // ignore existing clipping-mask paths
                if (isAxisRect(item)) {
                    rects.push(item);
                } else if (item.closed && item.pathPoints.length === 4) {
                    notes.push("A rotated or bent 4-point shape was skipped (only straight rectangles work).");
                }
                break;
        }
    }

    function overlap(a, b) {
        var w = Math.min(a[2], b[2]) - Math.max(a[0], b[0]);
        var h = Math.min(a[1], b[1]) - Math.max(a[3], b[3]);
        return (w > 0 && h > 0) ? w * h : 0;
    }

    // ---------- gather ----------

    for (var s = 0; s < sel.length; s++) collect(sel[s]);

    if (!rects.length || !images.length) {
        alert("I need at least one image AND at least one straight rectangle in the selection.\n\n" +
              "Found: " + images.length + " image(s), " + rects.length + " rectangle(s).");
        return;
    }

    // ---------- assign each rectangle to an image ----------
    // Each rectangle goes to the image it covers the most.
    // One image can collect several rectangles.

    var jobs = [];          // jobs[imageIndex] = { img, rects: [] }
    var lonelyRects = 0;

    for (var r = 0; r < rects.length; r++) {
        var rb = rects[r].geometricBounds;
        var best = -1, bestOverlap = 0;
        for (var m = 0; m < images.length; m++) {
            var ov = overlap(rb, images[m].geometricBounds);
            if (ov > bestOverlap) { bestOverlap = ov; best = m; }
        }
        if (best < 0) { lonelyRects++; continue; }
        if (!jobs[best]) jobs[best] = { img: images[best], rects: [] };
        jobs[best].rects.push(rects[r]);
    }

    if (lonelyRects > 0) notes.push(lonelyRects + " rectangle(s) weren't over any selected image and were left alone.");

    // ---------- crop one piece ----------
    // Copies the image, crops the copy to the rectangle, deletes the rectangle.

    function cropPiece(img, rect) {
        var cropBounds = rect.geometricBounds; // stroke width is ignored = exact path size
        var ib = img.geometricBounds;

        // Does the rectangle stay fully inside the image?
        var inside = cropBounds[0] >= ib[0] - TOL && cropBounds[2] <= ib[2] + TOL &&
                     cropBounds[1] <= ib[1] + TOL && cropBounds[3] >= ib[3] - TOL;

        var opts = new RasterizeOptions();
        opts.resolution = RESOLUTION;
        opts.antiAliasingMethod = AntiAliasingMethod.None;           // crisp edges, no soft fringe
        opts.colorModel = RasterizationColorModel.DEFAULTCOLORMODEL; // keep document colour mode
        opts.transparency = !inside;   // only add transparency if the box sticks out past the image
        opts.backgroundBlack = false;
        opts.clippingMask = false;
        opts.padding = 0;
        opts.convertSpotColors = false;
        opts.includeLayers = false;

        var oldName = "";
        try { oldName = img.name; } catch (e) {}

        var copy = img.duplicate(img, ElementPlacement.PLACEBEFORE);
        var newImg;
        try {
            newImg = doc.rasterize(copy, cropBounds, opts);
        } catch (e) {
            try { copy.remove(); } catch (e2) {}
            throw e;
        }

        // Illustrator normally removes the copy itself; make sure of it
        try { copy.typename; copy.remove(); } catch (e) {}

        if (oldName) { try { newImg.name = oldName; } catch (e) {} }
        if (!inside) notes.push("A rectangle reached past the edge of its image; that area was left transparent.");

        rect.remove();
        return newImg;
    }

    // ---------- crop everything ----------

    var results = [];
    var failed = 0;

    for (var j = 0; j < images.length; j++) {
        var job = jobs[j];
        if (!job) continue; // image had no rectangle on it: left untouched

        var allOk = true;
        for (var k = 0; k < job.rects.length; k++) {
            try {
                results.push(cropPiece(job.img, job.rects[k]));
            } catch (err) {
                allOk = false;
                failed++;
                notes.push("One crop failed: " + err.message);
            }
        }

        // Delete the original only if all its pieces were made successfully
        if (allOk) {
            try { job.img.remove(); } catch (e) {}
        } else {
            notes.push("An original image was kept because one of its crops failed.");
        }
    }

    // ---------- finish ----------

    doc.selection = null;
    for (var n = 0; n < results.length; n++) {
        try { results[n].selected = true; } catch (e) {}
    }

    var msg = "Made " + results.length + " cropped piece(s) at " + RESOLUTION + " PPI.";
    if (failed) msg += "\n" + failed + " failed.";
    if (notes.length) {
        // remove duplicate notes
        var seen = {}, uniq = [];
        for (var q = 0; q < notes.length; q++) {
            if (!seen[notes[q]]) { seen[notes[q]] = true; uniq.push(notes[q]); }
        }
        msg += "\n\n- " + uniq.join("\n- ");
    }
    alert(msg);

})();
