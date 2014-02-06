String.prototype.chunk = function (size) {
    //return [].concat.apply([], this.split('').map(function (x, i) { return i % size ? [] : this.slice(i, i + size) }, this))
    return this.match(/.{1,1000}/g);
}

String.prototype.trim = function () {
    return this.replace(/^\s*/, "").replace(/\s*$/, "");
}

