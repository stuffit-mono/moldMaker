/* 1) Create an instance of CSInterface. */
var csInterface = new CSInterface();

/* 2) Make a reference to your HTML button and add a click handler. */
var createButton = document.querySelector("#createBtn");
createButton.addEventListener("click", drawMold);

/* 3) Write a helper function to pass instructions to the ExtendScript side. */


// var coverResult;
// if (covered.checked ==true) {
//     coverResult = true;
// } else {
//     coverResult = false;
// }



// function drawMold() {
    //     csInterface.evalScript('processTest()');
    // }
function drawMold() {
    var depthLength = document.querySelector("#inputDepthLength").value;
    var widthFinger = document.querySelector("#inputWidthFinger").value;
    var heightFinger = document.querySelector("#inputHeightFinger").value;
    var depthFinger = document.querySelector("#inputDepthFinger").value;
    var thickness = document.querySelector("#inputThickness").value;
    var gap = document.querySelector("#inputGap").value;
    var margin = document.querySelector("#inputMargin").value;
    var covered = document.querySelector("#inputCover").checked;

    var mold_value = {
        userDepth: depthLength,
        userWidthFinger: widthFinger,
        userHeightFinger: heightFinger,
        userDepthFinger: depthFinger,
        userThickness: thickness,
        userGap: gap,
        userMargin: margin,
        userCovered: covered,
        
    }

    csInterface.evalScript('process(' + JSON.stringify(mold_value) + ')');
}