var fs = require('fs');
var promise = require('promise');

var StreamReader = exports = function(readStream){
    var self = this;
    var pos = 0;
    self.read = function (start, end){
        return new promise(function(fulfill, reject){
            fs.
        });
    }

    self.peek = function (start, end){

    }

    self.pos = function(){
        return pos;
    }

    self.hasNext = function(){
        
    }

    self.next = function(length){

    }
}

