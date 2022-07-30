// #include 'json2.min.js';

function _Log(contents) {
    $.writeln(contents);
}

// process();
// MEMO : UI 설정 종료
function process(obj) {
    _Log("프로세스 실행");

    var myDoc = app.activeDocument;
    var newLayerName = "Scripting Mold Layer";
    var layersDeleted = 0;
    var layerCount = myDoc.layers.length;
    var userValue = obj;

    var appSelection = myDoc.selection;
    if (appSelection.length==0 || appSelection.length > 1){
      alert("몰드 대상이 없거나 2개 이상입니다.")
      return false;
    } 

    var objectMarginMm = Number(userValue.userMargin);
    var gapMm = Number(userValue.userGap);
    var moldThicknessMm = Number(userValue.userThickness);
    var moldDepthMm = Number(userValue.userDepth);
    var cover = userValue.userCovered;
    var fingerWidthNum = Number(userValue.userWidthFinger);
    var fingerHeightNum = Number(userValue.userHeightFinger);
    var fingerSideNum = Number(userValue.userDepthFinger);

        


    var objectWidthPoints = appSelection[0].width;
    var objectHeightPoints = appSelection[0].height;
    var objectWidthMm = pointToMm(objectWidthPoints);
    var objectHeightMm = pointToMm(objectHeightPoints);
    var objectMarginPoints = mmToPoint(objectMarginMm);
    var gapPoints = mmToPoint(gapMm);
    var moldThicknessPoints = mmToPoint(moldThicknessMm)+mmToPoint(gapMm);
    var moldDepthPoints = mmToPoint(moldDepthMm);

    var moldWidth = objectWidthPoints + objectMarginPoints * 2;
    var moldHeight = objectHeightPoints + objectMarginPoints * 2;
    var selectedMoldRect = appSelection[0];
    // selectedMoldRect.locked = true;
    myDoc.selection = null;

    // var partsGap = moldThicknessPoints + gapPoints;



    var targetColor = new CMYKColor();
    targetColor.black = 50;
    targetColor.cyan = 0;
    targetColor.magenta = 0;
    targetColor.yellow = 0;
    var baseColor = new CMYKColor();
    baseColor.black = 0;
    baseColor.cyan = 0;
    baseColor.magenta = 100;
    baseColor.yellow = 0;
    var sideColor = new CMYKColor();
    sideColor.black = 0;
    sideColor.cyan = 100;
    sideColor.magenta = 0;
    sideColor.yellow = 0;
    var sizeColor = new CMYKColor();
    sizeColor.black = 70;
    sizeColor.cyan = 0;
    sizeColor.magenta = 0;
    sizeColor.yellow = 0;
    var arrowColor = new CMYKColor();
    arrowColor.black = 0;
    arrowColor.cyan = 0;
    arrowColor.magenta = 50;
    arrowColor.yellow = 100;

    // 레이어 초기화
    // for (var i = layerCount - 1; i >= 0; i--) {
    //     var targetLayer = myDoc.layers[i];
    //     var layerName = new String(targetLayer.name);
    //     if (layerName.indexOf(newLayerName) == 0) {
    //         myDoc.layers[i].remove();
    //         layersDeleted++;
    //     }
    // }

    // 레이어 생성
    var moldLayer = myDoc.layers.add();
    moldLayer.name = newLayerName;

    function mmToPoint(mm) {
        return mm * 2.834645;
    }
    function pointToMm(point) {
        return point / 2.834645;
    }

    function reverseStr(str) {
        var result = "";
        for (var i = str.length - 1; i >= 0; i--) {
            result = result + str[i];
        }
        return result;
    }

    if (fingerWidthNum % 2 == 0) {
        alert("핑거 숫자는 홀수여야 합니다");
    }

    // 선택항목 병합 함수
    function selectedMerge() {
        app.executeMenuCommand("group");
        app.executeMenuCommand("Live Pathfinder Add");
        app.executeMenuCommand("expandStyle");
        app.executeMenuCommand("ungroup");
        myDoc.selection = null;
    }

    // 몰드 타겟 생성(마진 포함)
    //     var objectWidthPoints = appSelection[0].width;
    // var objectHeightPoints = appSelection[0].height;
    // var objectWidthMm = pointToMm(objectWidthPoints);
    // var objectHeightMm = pointToMm(objectHeightPoints);
    // var objectMarginPoints = mmToPoint(objectMarginMm);
    // var gapPoints = mmToPoint(gapMm);
    // var moldThicknessPoints = mmToPoint(moldThicknessMm)+mmToPoint(gapMm);
    // var moldDepthPoints = mmToPoint(moldDepthMm);

    // var moldWidth = objectWidthPoints + objectMarginPoints * 2;
    // var moldHeight = objectHeightPoints + objectMarginPoints * 2;
        // var selectedMoldRect = appSelection[0];

    var targetTop = selectedMoldRect.top + objectMarginPoints;
    var targetBottom = selectedMoldRect.top - selectedMoldRect.height - objectMarginPoints;
    var targetLeft = selectedMoldRect.left - objectMarginPoints;
    var targetRight = selectedMoldRect.left + selectedMoldRect.width + objectMarginPoints;
    var baseDepth = moldDepthPoints;



    // 몰드 타겟 사각형 출력
    var targetLine = myDoc.pathItems.add();
    var targetLinePoint = [
        [targetLeft, targetTop],
        [targetRight, targetTop],
        [targetRight, targetBottom],
        [targetLeft, targetBottom],
    ];
    var developerCode = reverseStr("gnuoyeaJoeS yb gnitpircS");
    targetLine.stroked = true;
    targetLine.strokeColor = targetColor;
    targetLine.filled = false;
    targetLine.setEntirePath(targetLinePoint);
    targetLine.strokeDashes = [3];
    targetLine.closed = true;

    // pathItems 생성
    var baseLine = myDoc.pathItems.add();
    var topLine = myDoc.pathItems.add();
    var rightLine = myDoc.pathItems.add();
    var leftLine = myDoc.pathItems.add();
    var bottomLine = myDoc.pathItems.add();

    // 몰드 오브젝트 생성
    var moldLines = [
        {
            name: "baseLine",
            item: baseLine,
            line: null,
            points: [[], [], [], []],
            drawLines: [],
            lineOrder: [0, 1, 2, 3],
            strokeColor: baseColor,
            startPointX: targetLeft - moldThicknessPoints,
            startPointY: targetBottom - moldThicknessPoints,
        },
        {
            name: "rightLine",
            item: rightLine,
            line: null,
            points: [[], [], [], []],
            drawLines: [],
            lineOrder: [0, 1, 2, 3],
            strokeColor: sideColor,
            startPointX: targetLeft + moldWidth + moldThicknessPoints,
            startPointY: targetBottom,
            textPointX: targetRight + moldDepthPoints + moldThicknessPoints,
            textPointY: targetTop - moldHeight / 2,
        },
        {
            name: "topLine",
            item: topLine,
            line: null,
            points: [[], [], [], []],
            drawLines: [],
            lineOrder: [1, 2, 3, 0],
            strokeColor: sideColor,
            startPointX: targetLeft - moldThicknessPoints,
            startPointY: targetTop + moldThicknessPoints,
            textPointX: targetLeft + moldWidth / 2,
            textPointY:
                targetTop + moldDepthPoints + moldThicknessPoints * 5,
        },
        {
            name: "leftLine",
            item: leftLine,
            line: null,
            points: [[], [], [], []],
            drawLines: [],
            lineOrder: [2, 3, 0, 1],
            strokeColor: sideColor,
            startPointX: targetLeft - moldThicknessPoints - baseDepth,
            startPointY: targetBottom,
            textPointX:
                targetLeft - moldDepthPoints - moldThicknessPoints * 6,
            textPointY: targetTop - moldHeight / 2,
        },
        {
            name: "bottomLine",
            item: bottomLine,
            line: null,
            points: [[], [], [], []],
            drawLines: [],
            lineOrder: [3, 0, 1, 2],
            strokeColor: sideColor,
            startPointX: targetLeft - moldThicknessPoints,
            startPointY: targetBottom - moldThicknessPoints - baseDepth,
            textPointX: targetLeft + moldWidth / 2,
            textPointY:
                targetBottom - moldDepthPoints - moldThicknessPoints * 4,
        },
    ];

    var baseWidthFingerSize = moldWidth / fingerWidthNum;
    var baseHeightFingerSize = moldHeight / fingerHeightNum;
    var sideFingerSize = moldDepthPoints / fingerSideNum;
    var pointX = 0;
    var pointY = 0;

    // 포인트 좌표 추가
    function addPoint(line, num) {
        line.points[num].push([pointX, pointY]);
    }

    // 몰드 생성
    function printMold(obj) {
        drawLines(obj);
        obj.item.stroked = true;
        obj.item.strokeColor = obj.strokeColor;
        obj.item.strokeWidth = 1;
        obj.item.strokeDashes = [];
        obj.item.filled = false;
        obj.item.setEntirePath(obj.drawLines);
        obj.item.closed = false;
        obj.item.selected = true;
        obj.item.name = developerCode;
    }

    // 몰드라인 포인트 병합
    function drawLines(obj) {
        var pointlength = 0;
        for (var i = 0; i < obj.points.length; i++) {
            obj.drawLines = obj.drawLines.concat(
                obj.points[obj.lineOrder[i]]
            );
            pointlength = pointlength + obj.points[obj.lineOrder[i]].length;
        }
        var addStartPoint = obj.drawLines[obj.drawLines.length - 1];
        obj.drawLines.unshift(addStartPoint);
        var deleteLine = obj.points[obj.lineOrder[3]].length;
        // 갭이 0일때 베이스라인과 중첩된 라인 삭제
        if (gapPoints == 0 && obj.name != "baseLine") {
            obj.drawLines.splice(pointlength - deleteLine + 1);
        }
        // 커버가 false일때 상단 핑거 조인트 삭제
        if (cover == false && obj.name != "baseLine") {
            switch (obj.name) {
                case "rightLine":
                    obj.drawLines.splice(
                        fingerSideNum * 2 + 1,
                        fingerWidthNum * 2 - 2
                    );
                    break;
                case "topLine":
                    obj.drawLines.splice(
                        fingerSideNum * 2,
                        fingerWidthNum * 2 - 2
                    );
                    break;

                case "leftLine":
                    obj.drawLines.splice(
                        fingerSideNum * 2,
                        fingerWidthNum * 2 - 1
                    );
                    break;

                case "bottomLine":
                    obj.drawLines.splice(
                        fingerSideNum * 2,
                        fingerWidthNum * 2 - 1
                    );
                    break;
            }
        }
    }

    // 가로라인
    function baseDrawU(moldLine, moldNum, lineNum) {
        var fingerSize = sideFingerSize;
        var thickness = moldThicknessPoints;
        var lineGap = gapPoints;
        var nowFingerWidthNum = fingerWidthNum;
        if (moldNum != 0 && moldNum % 2 == 1) {
            nowFingerWidthNum = fingerSideNum;
        }
        if (moldNum % 2 == 0) { // 파츠가 베이스와 좌우일때,
            fingerSize = baseWidthFingerSize;
        }
        if (lineNum == 2) { // 상단 라인 그릴때 수치 반전 
            fingerSize = fingerSize * -1;
            thickness = thickness * -1;
            lineGap = gapPoints * -1;
        }
        for (var k = 0; k < nowFingerWidthNum; k++) {
            if (k % 2 == 0) { // 직선 그리기
                if ((k == 0 && moldNum % 2 == 0) || (k == nowFingerWidthNum - 1 && moldNum % 2 == 0)) {
                    if (moldNum == 0) {
                        pointX = pointX + fingerSize + thickness + lineGap;
                    } else {
                        pointX = pointX + fingerSize + thickness;
                    }
                } else if (
                    k == 0 ||
                    (k == nowFingerWidthNum - 1 && moldNum % 2 == 1)
                ) {
                    pointX = pointX + fingerSize - lineGap;
                } else {
                    if (moldNum == 0) {
                        pointX = pointX + fingerSize + lineGap * 2;
                    } else if (moldNum % 2 == 1) {
                        pointX = pointX + fingerSize - lineGap * 2;
                    } else {
                        pointX = pointX + fingerSize;
                    }
                }
                addPoint(moldLine, lineNum);
            } else { // 돌출 그리기 
                if (lineNum % 2 == 0 && moldNum != 0 && moldNum % 2 == 0) {
                    pointY = pointY - thickness;
                } else if (moldNum % 2 == 1) {
                    pointY = pointY - thickness;
                } else {
                    pointY = pointY + thickness;
                }
                addPoint(moldLine, lineNum);
                if (moldNum == 0) {
                    pointX = pointX + fingerSize - lineGap * 2;
                } else if (moldNum == 1 || moldNum == 3) {
                    pointX = pointX + fingerSize + lineGap * 2;
                } else {
                    pointX = pointX + fingerSize;
                }
                addPoint(moldLine, lineNum);
                if (lineNum % 2 == 0 && moldNum != 0 && moldNum % 2 == 0) {
                    pointY = pointY + thickness;
                } else if (moldNum % 2 == 1) {
                    pointY = pointY + thickness;
                } else {
                    pointY = pointY - thickness;
                }
                addPoint(moldLine, lineNum);
            }
        }
    }

    // 세로라인
    function baseDrawV(moldLine, moldNum, lineNum) {
        var fingerSize = baseHeightFingerSize;
        var thickness = moldThicknessPoints;
        var lineGap = gapPoints;
        var nowFingerHeightNum = fingerHeightNum;
        if (moldNum != 0 && moldNum % 2 == 0) {
            nowFingerHeightNum = fingerSideNum;
        }
        if (moldNum % 2 == 0 && moldNum != 0) {
            fingerSize = sideFingerSize;
        }
        if (lineNum == 3) {
            fingerSize = fingerSize * -1;
            thickness = thickness * -1;
            lineGap = gapPoints * -1;
        }
        for (var k = 0; k < nowFingerHeightNum; k++) {
            if (k % 2 == 0) {
                if (k == 0 || k == nowFingerHeightNum - 1) {
                    if (moldNum == 0) {
                        pointY = pointY + fingerSize + thickness + lineGap;
                    } else {
                        pointY = pointY + fingerSize;
                    }
                } else {
                    if (moldNum == 0) {
                        pointY = pointY + fingerSize + lineGap * 2;
                    } else {
                        pointY = pointY + fingerSize;
                    }
                }
                addPoint(moldLine, lineNum);
            } else {
                if (moldNum % 2 == 1) {
                    pointX = pointX + thickness;
                } else {
                    pointX = pointX - thickness;
                }
                addPoint(moldLine, lineNum);
                if (moldNum == 0) {
                    pointY = pointY + fingerSize - lineGap * 2;
                } else {
                    pointY = pointY + fingerSize;
                }
                addPoint(moldLine, lineNum);
                if (moldNum % 2 == 1) {
                    pointX = pointX - thickness;
                } else {
                    pointX = pointX + thickness;
                }
                addPoint(moldLine, lineNum);
            }
        }
    }

    // 전체몰드 라인 그리기
    function drawMoldLine(moldLines) {
        // 몰드라인 객체 참조하여그리기
        for (var i = 0; i < moldLines.length; i++) {
            // for (var i = 0; i < 1; i++) {
            pointX = moldLines[i].startPointX; // 몰드 시작점 초기화
            pointY = moldLines[i].startPointY;
            addPoint(moldLines[i], 0);
            // 개별 몰드라인 그리기
            for (var j = 0; j < 4; j++) {
                if (j % 2 == 0) {
                    // 몰드라인이 가로 일때
                    baseDrawU(moldLines[i], i, j);
                } else {
                    // 몰드라인이 세로일때
                    baseDrawV(moldLines[i], i, j);
                }
            }
            printMold(moldLines[i]);
        }
    }
    drawMoldLine(moldLines);

    // 커버 생성
    if (cover) {
        var coverLine = myDoc.pathItems.add();
        coverLine.stroked = true;
        coverLine.strokeColor = baseColor;
        coverLine.strokeWidth = 1;
        coverLine.strokeDashes = [];
        coverLine.filled = false;
        coverLine.closed = false;
        coverLine.selected = true;
        coverLine.points = moldLines[0].points;
        coverLine.lineOrder = [2, 3, 0, 1];
        coverLine.drawLines = [];
        var pointlength = 0;

        for (var i = 0; i < coverLine.points.length; i++) {
            coverLine.drawLines = coverLine.drawLines.concat(
                coverLine.points[coverLine.lineOrder[i]]
            );
            pointlength = pointlength + coverLine.points[coverLine.lineOrder[i]].length;
        }
        var addStartPoint = coverLine.drawLines[coverLine.drawLines.length - 1];
        coverLine.drawLines.unshift(addStartPoint);
        var deleteLine = coverLine.points[coverLine.lineOrder[3]].length;
        // 갭이 0일때 베이스라인과 중첩된 라인 삭제
        if (gapPoints == 0) {
            coverLine.drawLines.splice(pointlength - deleteLine + 1);
            coverLine.drawLines.unshift(moldLines[1].points[2][moldLines[1].points[2].length - 1]);
            coverLine.drawLines.push(moldLines[1].points[0][0]);
        }
            
        coverLine.setEntirePath(coverLine.drawLines);

        if (gapPoints > 0) {
            coverLine.translate(-baseLine.width - leftLine.width, 0);
        } else {
            coverLine.translate(-baseLine.width - leftLine.width + moldThicknessPoints, 0);
        }
    }

    // 몰드 이동
    if (gapPoints > 0) {
        topLine.translate(0, moldThicknessPoints);
        rightLine.translate(moldThicknessPoints, 0);
        leftLine.translate(-moldThicknessPoints, 0);
        bottomLine.translate(0, -moldThicknessPoints);
    }

    app.executeMenuCommand("group");
    var moldLineGroup = myDoc.selection[0];
    myDoc.selection = null;

    // 폰트 설정
    var fonts = app.textFonts;
    // var targetFont = "Arial-BoldMT";
    var targetFont = "RobotoMono-Regular";
    var fontsName = "";
    var useFont;

    for (var i = 0; i < fonts.length; i++) {
        if (fonts[i].name == targetFont) {
            useFont = fonts[i];
            break;
        }
    }

    // 치수 출력
    // 전체 가로 치수
    var lineTale = 10;
    var lineTextGap = 3;

    var originBoxWidth = myDoc.textFrames.add();
    var originBoxHeight = myDoc.textFrames.add();
    var totalWidth = myDoc.textFrames.add();
    var totalHeight = myDoc.textFrames.add();
    var boxWidth = myDoc.textFrames.add();
    var boxHeight = myDoc.textFrames.add();
    var boxDepth = myDoc.textFrames.add();
    var gapText = myDoc.textFrames.add();
    var moldThicknessText = myDoc.textFrames.add();
    var originBoxWidth_ArrowLine_Left = myDoc.pathItems.add();
    var originBoxWidth_ArrowLine_Right = myDoc.pathItems.add();
    var originBoxHeight_ArrowLine_Top = myDoc.pathItems.add();
    var originBoxHeight_ArrowLine_Bottom = myDoc.pathItems.add();
    var totalWidthLine_Left = myDoc.pathItems.add();
    var totalWidthLine_Right = myDoc.pathItems.add();
    var boxWidthLine_Left = myDoc.pathItems.add();
    var boxWidthLine_Right = myDoc.pathItems.add();
    var totalHeightLine_Top = myDoc.pathItems.add();
    var totalHeightLine_Bottom = myDoc.pathItems.add();
    var boxHeightLine_Top = myDoc.pathItems.add();
    var boxHeightLine_Bottom = myDoc.pathItems.add();
    var boxDepth_Left = myDoc.pathItems.add();
    var boxDepth_Right = myDoc.pathItems.add();

    // 텍스트 공통 설정
    function sizeTextsSetting(textFrame) {
        textFrame.textRange.characterAttributes.size = 10;
        textFrame.textRange.characterAttributes.fillColor = sizeColor;
        textFrame.textRange.characterAttributes.textFont = useFont;
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
        textFrame.selected = true;
    }

    // 내부 치수 정보

    originBoxWidth.contents = objectWidthMm.toFixed(2) + "mm";
    sizeTextsSetting(originBoxWidth);
    originBoxWidth.left = selectedMoldRect.left + selectedMoldRect.width / 2 - originBoxWidth.width / 2;
    originBoxWidth.top = selectedMoldRect.top - selectedMoldRect.height / 2 + originBoxWidth.height /2;

    originBoxHeight.contents = objectHeightMm.toFixed(2) + "mm";
    sizeTextsSetting(originBoxHeight);
    originBoxHeight.left = selectedMoldRect.left + (selectedMoldRect.width / 4)- originBoxHeight.width / 2;
    originBoxHeight.top = selectedMoldRect.top - selectedMoldRect.height / 2+ originBoxHeight.width;
    originBoxHeight.rotate(90)
    
    gapText.contents = "margin: "+ objectMarginMm + "mm";
    sizeTextsSetting(gapText);
    gapText.textRange.characterAttributes.size = 10;
    gapText.left = selectedMoldRect.left + selectedMoldRect.width / 2 - gapText.width / 2;
    gapText.top = targetLine.top - (targetLine.top - selectedMoldRect.top)/2 + gapText.height /2;
    
    moldThicknessText.contents = "Thickness: " + moldThicknessMm + "mm \n\nGap: " + gapMm +"mm";
    sizeTextsSetting(moldThicknessText);
    moldThicknessText.textRange.characterAttributes.size = 10;
    moldThicknessText.left = selectedMoldRect.left + selectedMoldRect.width / 2 - moldThicknessText.width / 2;
    moldThicknessText.top = selectedMoldRect.top-selectedMoldRect.height/4 + moldThicknessText.height /2;

    
    // 외부 치수 정보

    totalWidth.contents = pointToMm(moldLineGroup.width).toFixed(2) + "mm";
    sizeTextsSetting(totalWidth);
    totalWidth.left = bottomLine.left + bottomLine.width / 2 - totalWidth.width/2;
    totalWidth.top = bottomLine.top - bottomLine.height - lineTale*2;

    totalHeight.contents = pointToMm(moldLineGroup.height).toFixed(2) + "mm";
    sizeTextsSetting(totalHeight);
    totalHeight.rotate(-90);
    totalHeight.left = rightLine.left+rightLine.width-totalHeight.width/2+lineTale*2;
    totalHeight.top = rightLine.top- rightLine.height/2 + totalHeight.height / 2;

    boxWidth.contents = pointToMm(baseLine.width).toFixed(2) + "mm";
    sizeTextsSetting(boxWidth);
    boxWidth.left = topLine.left + topLine.width / 2 - boxWidth.width / 2;
    boxWidth.top = topLine.top + boxWidth.height/2+lineTale*2;

    boxHeight.contents = pointToMm(baseLine.height).toFixed(2) + "mm";
    sizeTextsSetting(boxHeight);
    boxHeight.left = moldLineGroup.left - boxHeight.width/2 - lineTale*2;
    boxHeight.top = leftLine.top- leftLine.height/2 + boxHeight.height / 2;
    boxHeight.rotate(90);

    boxDepth.contents = moldDepthMm.toFixed(2) + "mm";
    sizeTextsSetting(boxDepth);
    var sideTextPosition = 0;
    if (cover == false && gapPoints > 0) {
        sideTextPosition = moldThicknessPoints/2
    } else if (cover == true && gapPoints == 0) {
        sideTextPosition = -moldThicknessPoints/2
    }
    boxDepth.left = rightLine.left + rightLine.width / 2 - boxDepth.width / 2 + sideTextPosition;
    boxDepth.top = rightLine.top + moldThicknessPoints * 3;

    // 라인 공통 설정
    function sizeLineSetting(sizeLine) {
        sizeLine.name = sizeLine;
        sizeLine.stroked = true;
        sizeLine.strokeWidth = 0.7;
        sizeLine.strokeColor = sizeColor;
        sizeLine.filled = false;
        sizeLine.closed = false;
        sizeLine.strokeDashes = [4, 2];
        sizeLine.selected = true;
    }

    sizeLineSetting(originBoxWidth_ArrowLine_Left);
    sizeLineSetting(originBoxWidth_ArrowLine_Right);
    sizeLineSetting(originBoxHeight_ArrowLine_Top);
    sizeLineSetting(originBoxHeight_ArrowLine_Bottom);
    sizeLineSetting(totalWidthLine_Left);
    sizeLineSetting(totalWidthLine_Right);
    sizeLineSetting(totalHeightLine_Top);
    sizeLineSetting(totalHeightLine_Bottom);
    sizeLineSetting(boxWidthLine_Left);
    sizeLineSetting(boxWidthLine_Right);
    sizeLineSetting(boxHeightLine_Top);
    sizeLineSetting(boxHeightLine_Bottom);
    sizeLineSetting(boxDepth_Left);
    sizeLineSetting(boxDepth_Right);

    // 라인 개별 설정
  

    originBoxWidth_ArrowLine_Left.setEntirePath([
        [selectedMoldRect.left, moldLines[1].textPointY],
        [originBoxWidth.left-lineTextGap, moldLines[1].textPointY],
    ]);
    originBoxWidth_ArrowLine_Right.setEntirePath([
        [selectedMoldRect.left+selectedMoldRect.width, moldLines[1].textPointY],
        [originBoxWidth.left+originBoxWidth.width+lineTextGap, moldLines[1].textPointY],
    ]);
    arrowHead(originBoxWidth_ArrowLine_Left, originBoxWidth_ArrowLine_Right);
    
    originBoxHeight_ArrowLine_Top.setEntirePath([
        [selectedMoldRect.left+(selectedMoldRect.width/4), selectedMoldRect.top],
        [selectedMoldRect.left+(selectedMoldRect.width/4), originBoxHeight.top+lineTextGap],
    ]);
    originBoxHeight_ArrowLine_Bottom.setEntirePath([
        [selectedMoldRect.left+(selectedMoldRect.width/4), selectedMoldRect.top - selectedMoldRect.height],
        [selectedMoldRect.left+(selectedMoldRect.width/4), originBoxHeight.top-originBoxHeight.height-lineTextGap],
    ]);
    arrowHead(originBoxHeight_ArrowLine_Top, originBoxHeight_ArrowLine_Bottom);
    originBoxWidth_ArrowLine_Left.strokeColor = arrowColor;
    originBoxWidth_ArrowLine_Right.strokeColor = arrowColor;
    originBoxHeight_ArrowLine_Top.strokeColor = arrowColor;
    originBoxHeight_ArrowLine_Bottom.strokeColor = arrowColor;
    

    totalWidthLine_Left.setEntirePath([
        [moldLineGroup.left, totalWidth.top-totalWidth.height/2 + lineTale],
        [moldLineGroup.left, totalWidth.top-totalWidth.height/2],
        [totalWidth.left-lineTextGap, totalWidth.top-totalWidth.height/2],
    ]);

    totalWidthLine_Right.setEntirePath([
        [ totalWidth.left + totalWidth.width +lineTextGap, totalWidth.top-totalWidth.height/2],
        [ moldLineGroup.left + moldLineGroup.width, totalWidth.top-totalWidth.height/2],
        [moldLineGroup.left + moldLineGroup.width, totalWidth.top-totalWidth.height/2 + lineTale],
    ]);

    totalHeightLine_Top.setEntirePath([
        [ totalHeight.left+totalHeight.width/2-lineTale, moldLineGroup.top],
        [ totalHeight.left+totalHeight.width/2, moldLineGroup.top],
        [ totalHeight.left+totalHeight.width/2, totalHeight.top+lineTextGap],
    ]);

    totalHeightLine_Bottom.setEntirePath([
        [ totalHeight.left+totalHeight.width/2, totalHeight.top-totalHeight.height-lineTextGap],
        [ totalHeight.left+totalHeight.width/2, moldLineGroup.top - moldLineGroup.height],
        [ totalHeight.left+totalHeight.width/2 - lineTale, moldLineGroup.top - moldLineGroup.height],
    ]);

    boxWidthLine_Left.setEntirePath([
        [baseLine.left, boxWidth.top - boxWidth.height / 2 - lineTale],
        [baseLine.left, boxWidth.top - boxWidth.height / 2],
        [boxWidth.left- lineTextGap, boxWidth.top - boxWidth.height / 2],
    ]);

    boxWidthLine_Right.setEntirePath([
        [boxWidth.left+ boxWidth.width+lineTextGap, boxWidth.top - boxWidth.height / 2],
        [baseLine.left+ baseLine.width, boxWidth.top - boxWidth.height / 2],
        [baseLine.left+ baseLine.width, boxWidth.top - boxWidth.height / 2-lineTale],
    ]);

    boxHeightLine_Top.setEntirePath([
        [ boxHeight.left + boxHeight.width/2 + lineTale, baseLine.top ],
        [ boxHeight.left + boxHeight.width/2, baseLine.top, ],
        [ boxHeight.left + boxHeight.width/2, boxHeight.top + lineTextGap],
    ]);

    boxHeightLine_Bottom.setEntirePath([
        [ boxHeight.left + boxHeight.width/2, boxHeight.top - boxHeight.height - lineTextGap ],
        [ boxHeight.left + boxHeight.width/2, baseLine.top - baseLine.height ],
        [ boxHeight.left + boxHeight.width/2+lineTale, baseLine.top - baseLine.height ],
    ]);

    if (gapMm == 0) {
        boxDepth_Left.setEntirePath([
            [rightLine.left, boxDepth.top - boxDepth.height / 2 - lineTale],
            [rightLine.left, boxDepth.top - boxDepth.height / 2],
            [boxDepth.left - lineTextGap, boxDepth.top - boxDepth.height / 2],
        ]);
        if (cover == false) {
            boxDepth_Right.setEntirePath([
                [ boxDepth.left + boxDepth.width + lineTextGap, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width, boxDepth.top - boxDepth.height / 2 - lineTale ],
            ]);
        } else {
            boxDepth_Right.setEntirePath([
                [ boxDepth.left + boxDepth.width + lineTextGap, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width - moldThicknessPoints, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width - moldThicknessPoints, boxDepth.top - boxDepth.height / 2 - lineTale ],
            ]);
        }
    } else {
        boxDepth_Left.setEntirePath([
            [ rightLine.left + moldThicknessPoints, boxDepth.top - boxDepth.height / 2 - lineTale],
            [ rightLine.left + moldThicknessPoints, boxDepth.top - boxDepth.height / 2 ],
            [boxDepth.left - lineTextGap, boxDepth.top - boxDepth.height / 2],
        ]);
        if (cover == false) {
            boxDepth_Right.setEntirePath([
                [ boxDepth.left + boxDepth.width + lineTextGap, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width, boxDepth.top - boxDepth.height / 2 ],
                [ rightLine.left + rightLine.width, boxDepth.top - boxDepth.height / 2 - lineTale],
            ]);
        } else {
            boxDepth_Right.setEntirePath([
                [ boxDepth.left + boxDepth.width + lineTextGap, boxDepth.top - boxDepth.height / 2],
                [ rightLine.left + rightLine.width - moldThicknessPoints, boxDepth.top - boxDepth.height / 2],
                [ rightLine.left + rightLine.width - moldThicknessPoints, boxDepth.top - boxDepth.height / 2 - lineTale],
            ]);
        }
    }

    function arrowHead(highLine, lowLine) {
        var arrowLine = [highLine, lowLine];
        var direction;
        direction = highLine.width > highLine.height ? "U" : "V"
        for (var i = 0; i < arrowLine.length; i++) {
            var arrowLength = 5;
            var arrowHeadPoint = arrowLine[i].pathPoints[0]
            arrowLine[i].arrowPoint = myDoc.pathItems.add()
            var arrowHead = arrowLine[i].arrowPoint;
            arrowHead.stroked = true;
            arrowHead.strokeWidth = 0.7;
            arrowHead.strokeColor = arrowColor;
            arrowHead.filled = false;
            arrowHead.closed = false;
            arrowHead.strokeDashes = [4, 2];
            arrowLength = i == 1 ? -arrowLength : arrowLength;
            if (direction == "U") {
                arrowHead.setEntirePath([
                    [arrowHeadPoint.anchor[0] + arrowLength, arrowHeadPoint.anchor[1] + arrowLength],
                    [arrowHeadPoint.anchor[0], arrowHeadPoint.anchor[1]],
                    [arrowHeadPoint.anchor[0] + arrowLength, arrowHeadPoint.anchor[1] - arrowLength],
                ]);
            } else {
                arrowLine[i].arrowPoint.setEntirePath([
                    [arrowHeadPoint.anchor[0] - arrowLength, arrowHeadPoint.anchor[1] - arrowLength],
                    [arrowHeadPoint.anchor[0], arrowHeadPoint.anchor[1]],
                    [arrowHeadPoint.anchor[0] + arrowLength, arrowHeadPoint.anchor[1] - arrowLength],
                ]);
            }
            arrowLine[i].arrowPoint.selected = true;
        }
    }

    var documentTextFrames = myDoc.textFrames;
    for (var i = 0; i < documentTextFrames.length; i++) {
        documentTextFrames[i].selected = true;
    }

    // selectedMoldRect.selected = true;
    targetLine.selected = true;

    var textLayer = myDoc.layers.add();
    textLayer.name = newLayerName + "_text";

    var textLayerItems = myDoc.selection;

    for (var i = 0; i < textLayerItems.length; i++) {
        textLayerItems[i].move(
            textLayer,
            ElementPlacement.PLACEATBEGINNING
        );
    }

    app.executeMenuCommand("group");
};

_Log("프로그램 종료")