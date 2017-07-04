var DxfParser = require('dxf-parser'),
    fs = require('fs'),
    path = require('path');

var O_PATH = path.join('data/cnc', process.argv[2] || ''),
    I_PATH = path.join('data/dxf', process.argv[2] || ''),
    OUTPUT_FILE_HOLZMA = "holzma.json",
    OUTPUT_FILE_DXF = "dxf.json";

makeDirSync(O_PATH);

fs.readdir(I_PATH, function (err, dxfFiles) {
    if (err) {
        throw err;
    }
    dxfFiles.map(function (dxfFile) {
        return path.join(I_PATH, dxfFile);
    }).filter(function (dxfFile) {
        return fs.statSync(dxfFile).isFile() && path.extname(dxfFile) == '.dxf';
    }).forEach(function (dxfFile) {
        biesseOutput(dxfFile);
    });
    outputHolzma();
});

function makeDirSync(dirPath) {
    try {
        fs.mkdirSync(dirPath);
    } catch (err) {
        if (err.code == 'EEXIST') {
            console.log(dirPath + ' already exist');
            process.exit();
        }
    }
}

function biesseOutput(dxfFile) {

    var outputFile = path.join(O_PATH, path.basename(dxfFile, path.extname(dxfFile))) + '.bpp';

    var fileStream = fs.createReadStream(dxfFile, {encoding: 'utf8'});

    var outputHeader = [],
        outputVariables = [],
        outputProgramm = [],
        outputVBScriptHeader = [],
        outputVBScript = [],
        tdcodes = {},
        outputTdcodes = [],
        side,           // Panel side
        tdcode,         // Tool code
        i,
        propName,
        entity,
        geomIdentificator = 1020;

    var startPointX,    // Starting point coordinate by x
        startPointY,    // Starting point coordinate by y
        endPointX,      // End point coordinate by x
        endPointY,      // End point coordinate by y
        curveRadius,    // Curve radius
        curveSolution,  // Curve solution type
        curveDirection, // Curve direction (CW or CCW)
        curveBulge,     // Curve bulge
        mlid,           // Milling id
        mldia,          // Milling tool diameter
        mldp;           // Milling depth

    var bhx,            // Horizontal boring coordinates
        bhdp,           // Horizontal boring depth
        bhdia,          // Horizontal boring diameter
        bhver,          // Vertex of the panel side
        bhid,           // Horizontal boring id
        sideIndex, bhxIndex, bhdpIndex, bhdiaIndex;

    var bvx,            // Vertical boring coordinate by x
        bvy,            // Vertical boring coordinate by y
        bvdp,           // Vertical boring depth
        bvdia,          // Vertical boring diameter
        bvver,          // Vertex of the panel side
        bvid,           // Vertical boring id
        bvthr,          // Vertical boring type (through = YES or NO)
        toolType;       // Tool type

    var parser = new DxfParser();

    parser.parseStream(fileStream, function (err, dxf) {
        if (err) return console.error(err.stack);
        fs.writeFileSync(OUTPUT_FILE_DXF, JSON.stringify(dxf, null, 3));

        // Add headers
        outputHeader.push('[HEADER]\nTYPE=BPP\nVER=120\n\n[DESCRIPTION]\n|');
        outputTdcodes.push('\n\n[TDCODES]\nVER=1\nCONVF=1.000000\nBUGFLAGS=3');
        outputVBScript.push('');

        for (i = 0; i < dxf.entities.length; i++) {
            entity = dxf.entities[i];
            for (propName in entity) {
                if (entity.hasOwnProperty(propName)) {
                    // Add VARIABLES and panel sizes
                    if (propName == 'layer' && entity[propName].indexOf('Z_') == 0) {
                        var lpx = 0, lpy = 0, lpz = 0;
                        for (var k = 0; k < entity['vertices'].length; k++) {
                            if (entity['vertices'][k].x > lpx) {
                                lpx = +entity['vertices'][k].x
                            }
                            if (entity['vertices'][k].y > lpy) {
                                lpy = +entity['vertices'][k].y
                            }
                            lpz = +entity['vertices'][k].layer.substring(2);
                        }
                        outputVariables.push('\n\n[VARIABLES]\nPAN=LPX|' + lpx + '||4|\nPAN=LPY|' + lpy + '||4|\nPAN=LPZ|' + lpz + '||4|\nPAN=ORLST|"1"||0|\nPAN=SIMMETRY|1||0|\nPAN=TLCHK|0||0|\nPAN=TOOLING|""||0|\nPAN=CUSTSTR|""||0|\nPAN=FCN|1.000000||0|\nPAN=XCUT|0||4|\nPAN=YCUT|0||4|\nPAN=JIGTH|0||4|\nPAN=CKOP|0||0|\nPAN=UNIQUE|0||0|\nPAN=MATERIAL|"wood"||0|\nPAN=PUTLST|""||0|\nPAN=OPPWKRS|0||0|\nPAN=UNICLAMP|0||0|\nPAN=CHKCOLL|0||0|\nPAN=WTPIANI|0||0|\nPAN=COLLTOOL|0||0|\n\n[PROGRAM]\n\n');
                        outputVBScriptHeader.push('\n\n[VBSCRIPT]\nOption Explicit\nDim mm: mm = 1.000000000000000000\nDim inc: inc = 25.399999999999999000\nDim SIDE: SIDE = -1\nDim LPX: LPX = ' + lpx + '\nDim LPY: LPY = ' + lpy + '\nDim LPZ: LPZ = ' + lpz + '\nDim ORLST: ORLST = "1"\nDim SIMMETRY: SIMMETRY = 1\nDim TLCHK: TLCHK = 0\nDim TOOLING: TOOLING = ""\nDim CUSTSTR: CUSTSTR = ""\nDim FCN: FCN = 1.000000\nDim XCUT: XCUT = 0\nDim YCUT: YCUT = 0\nDim JIGTH: JIGTH = 0\nDim CKOP: CKOP = 0\nDim UNIQUE: UNIQUE = 0\nDim MATERIAL: MATERIAL = "wood"\nDim PUTLST: PUTLST = ""\nDim OPPWKRS: OPPWKRS = 0\nDim UNICLAMP: UNICLAMP = 0\nDim CHKCOLL: CHKCOLL = 0\nDim WTPIANI: WTPIANI = 0\nDim COLLTOOL: COLLTOOL = 0\nSub Main()\nCall ProgBuilder.SetPanel(LPX*FCN, LPY*FCN, LPZ*FCN, ORLST, SIMMETRY, TLCHK, TOOLING, CUSTSTR, FCN, XCUT*FCN, YCUT*FCN, JIGTH*FCN, CKOP, UNIQUE, MATERIAL, PUTLST, OPPWKRS, UNICLAMP, CHKCOLL, WTPIANI, COLLTOOL)')
                    }
                }
            }
        }

        for (i = 0; i < dxf.entities.length; i++) {
            entity = dxf.entities[i];
            for (propName in entity) {
                if (entity.hasOwnProperty(propName)) {
                    // Add horizontal boring
                    if (propName == 'layer' && entity[propName].indexOf('BH_') == 0 && entity[propName].indexOf('_SIDE=') != -1) {
                        bhid = outputProgramm.length + 10000000;
                        sideIndex = entity[propName].indexOf('_SIDE=') + 6;
                        bhxIndex = entity[propName].indexOf('_BHX=') + 5;
                        bhdpIndex = entity[propName].indexOf('_BHDP=') + 6;
                        bhdiaIndex = entity[propName].indexOf('_BHDIA=') + 7;
                        side = entity[propName].substring(sideIndex, bhxIndex - 5);
                        bhx = +entity[propName].substring(bhxIndex, bhdpIndex - 6);
                        bhdp = +entity[propName].substring(bhdpIndex, bhdiaIndex - 7);
                        bhdia = +entity[propName].substring(bhdiaIndex);
                        tdcode = 'TDCODE' + '1' + String(bhdia * 100) + String(bhdp * 100);
                        if (side == 2 || side == 3) {
                            bhver = 1
                        }
                        if (side == 1 || side == 4) {
                            bhver = 4
                        }
                        tdcodes[tdcode] = '(MST)<LN=1,NJ=' + tdcode + ',TYW=1,NT=1,>\n(GEN)<WT=1,DL=1,>\n(TOO)<DI=' + bhdia + ',SP=8,CL=0,COD=----------,RO=-1,TY=0,ACT=0,NCT=,DCT=5.000000,TCT=0.000000,DICT=20.000000,DFCT=80.000000,PCT=60.000000,>\n(IO)<AI=0.000,AO=0.000,DA=0.000,DT=0.000,DD=0.000,IFD=0.00,OFD=0.00,IN=0,OUT=0,PR=0,ETCI=0,ITI=0,TLI=0.00,THI=0.00,ETCO=0,ITO=0,TLO=0.00,THO=0.00,PDI=0.00,PDO=0.00,>\n(WRK)<OP=1,CO=0,HH=0.000,DR=0,PV=0,PT=0,TC=-1,DP=' + bhdp + ',SM=0.000,TT=0,RC=0,BD=0,SW=0,IC="",IM="",IA="",PC=0,BL=0,PU=0,EA=0,EEA=0,SP=0,AP=0,>\n(SPD)<AF=0.00,CF=0.000,DS=0.00,FE=0.00,RT=0.00,OF=0.00,>\n(MOR)<PE=0.000000,TG=0.000000,TL=0.000000,WH=0.000000,>';
                        outputProgramm.push('@ BH, "' + tdcode + '", "", ' + bhid + ' : ' + side + ', "' + bhver + '", ' + bhx + ', 0, 0, ' + bhdp + ', ' + bhdia + ', 0, -1, 32, 32, 50, 0, 45, 0, "", 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, -1, "P1001", 0');
                        outputVBScript.push('SIDE = ' + side + ': Call ProgBuilder.AddBoring(0, ' + bhid + ', "' + tdcode + '"    , ' + side + ', "' + bhver + '", (' + bhx + ')*FCN, (0)*FCN, (0)*FCN, (' + bhdp + ')*FCN, (' + bhdia + ')*FCN, NO, rpNO, (32)*FCN, (32)*FCN, (50)*FCN, 0, 45, 0, "", YES, 0, 0, NO, azrNO, (0)*FCN, (0)*FCN, 0, (0)*FCN, YES, YES, NO, 0, NO, 0, -1, "P1001", (0)*FCN): SIDE = -1');
                    }
                    // Add vertical boring
                    if (propName == 'layer' && entity[propName].indexOf('BV_') == 0) {
                        bvid = outputProgramm.length + 10000000;
                        side = 0;
                        bvx = +entity['center']['x'];
                        bvy = +entity['center']['y'];
                        bvdia = +entity['radius'] * 2;
                        bvver = 2;
                        bvdp = +entity[propName].substring(entity[propName].indexOf('BV_') + 3);
                        if (bvdp >= lpz) {
                            toolType = 1;
                            bvdp = 5;
                            bvthr = 1;
                        } else {
                            toolType = bvdia < 15 ? 0 : 3;
                            bvthr = 0;
                        }
                        tdcode = 'TDCODE' + '2' + String(bvdia * 100) + String(bvdp * 100);
                        tdcodes[tdcode] = '(MST)<LN=1,NJ=' + tdcode + ',TYW=1,NT=1,>\n(GEN)<WT=1,DL=1,>\n(TOO)<DI=' + bvdia + ',SP=' + bvdia + ',CL=0,COD=----------,RO=-1,TY=' + toolType + ',ACT=0,NCT=,DCT=5.000000,TCT=0.000000,DICT=20.000000,DFCT=80.000000,PCT=60.000000,>\n(IO)<AI=0.000,AO=0.000,DA=0.000,DT=0.000,DD=0.000,IFD=0.00,OFD=0.00,IN=0,OUT=0,PR=0,ETCI=0,ITI=0,TLI=0.00,THI=0.00,ETCO=0,ITO=0,TLO=0.00,THO=0.00,PDI=0.00,PDO=0.00,>\n(WRK)<OP=1,CO=0,HH=0.000,DR=0,PV=0,PT=' + bvthr + ',TC=-1,DP=' + bvdp + ',SM=0.000,TT=0,RC=0,BD=0,SW=0,IC="",IM="",IA="",PC=0,BL=0,PU=0,EA=0,EEA=0,SP=0,AP=0,>\n(SPD)<AF=0.00,CF=0.000,DS=0.00,FE=0.00,RT=0.00,OF=0.00,>\n(MOR)<PE=0.000000,TG=0.000000,TL=0.000000,WH=0.000000,>';
                        outputProgramm.push('@ BV, "' + tdcode + '", "", ' + bvid + ' : ' + side + ', "' + bvver + '", ' + bvx + ', ' + bvy + ', 0, ' + bvdp + ', ' + bvdia + ', ' + bvthr + ', -1, 0, 0, 0, 0, 0, 0, "", 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, -1, "P1007", 0');
                        outputVBScript.push('SIDE = ' + side + ': Call ProgBuilder.AddBoring(0, ' + bvid + ', "' + tdcode + '"    , ' + side + ', "' + bvver + '", (' + bvx + ')*FCN, (' + bvy + ')*FCN, (0)*FCN, (' + bvdp + ')*FCN, (' + bvdia + ')*FCN, ' + (bvthr == 1 ? 'YES' : 'NO') + ', rpNO, (0)*FCN, (0)*FCN, (0)*FCN, 0, 0, 0, "", YES, 0, 0, NO, azrNO, (0)*FCN, (0)*FCN, 0, (0)*FCN, YES, NO, NO, 0, NO, 0, -1, "P1007", (0)*FCN): SIDE = -1');
                    }

                    // Add milling
                    if (entity[propName] == 'POLYLINE' && entity['layer'].indexOf('ML_') == 0) {
                        geomIdentificator++;
                        mldp = lpz + 2;
                        mldia = entity['layer'].substring(entity['layer'].indexOf('ML_') + 3, entity['layer'].indexOf('MM'));
                        startPointX = +entity.vertices[0]['x'];
                        startPointY = +entity.vertices[0]['y'];
                        tdcode = 'TDCODE' + '3' + String(mldia * 100) + String(mldp * 100);
                        mlid = outputProgramm.length + 10000000;
                        outputProgramm.push('@ ROUT, "' + tdcode + '", "", ' + mlid + ' : "' + geomIdentificator + '", 0, "2", 0, ' + mldp + ', "", 1, ' + mldia + ', -1, 0, 0, 32, 32, 50, 0, 45, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, -1, 0');
                        mlid = outputProgramm.length + 10000000;
                        outputProgramm.push('   @ START_POINT, "", "", ' + mlid + ' : ' + startPointX + ', ' + startPointY + ', 0');
                        tdcodes[tdcode] = '(MST)<LN=3,NJ=' + tdcode + ',TYW=2,NT=1,>\n(GEN)<WT=2,DL=0,>\n(TOO)<DI=' + mldia + ',SP=' + mldia + ',CL=1,COD=25NEW,RO=-1,TY=103,ACT=0,NCT=,DCT=5.000000,TCT=0.000000,DICT=20.000000,DFCT=80.000000,PCT=60.000000,>\n(IO)<AI=0.000,AO=0.000,DA=35.000,DT=35.000,DD=0.000,IFD=0.00,OFD=0.00,IN=2,OUT=2,PR=0,ETCI=0,ITI=0,TLI=0.00,THI=0.00,ETCO=0,ITO=0,TLO=0.00,THO=0.00,PDI=0.00,PDO=0.00,>\n(WRK)<OP=1,CO=0,HH=0.000,DR=0,PV=0,PT=0,TC=1,DP=' + mldp + ',SM=0,TT=0,RC=0,BD=0,SW=0,IC="",IM="",IA="",PC=0,BL=0,PU=0,EA=0,EEA=0,SP=0,AP=0,>\n(SPD)<AF=0.00,CF=0.000,DS=0.00,FE=0.00,RT=0.00,OF=0.00,>\n(MOR)<PE=0.000000,TG=0.000000,TL=0.000000,WH=0.000000,>'
                        outputVBScript.push('SIDE = 0: Call ProgBuilder.StartRout(0, ' + mlid + ', "' + tdcode + '"    , "' + geomIdentificator + '", 0, "2", (0)*FCN, (' + mldp + ')*FCN, "", YES, (' + mldia + ')*FCN, rpNO, (0)*FCN, (0)*FCN, (32)*FCN, (32)*FCN, (50)*FCN, 0, 45, YES, 0, 0, 0, (0)*FCN, (0)*FCN, azrNO, NO, NO, NO, 0, (0)*FCN, YES, NO, (0)*FCN, 0, NO, 0, (0)*FCN, 0, (0)*FCN, NO, (0)*FCN, YES, 0, -1, (0)*FCN)\nCall ProgBuilder.AddPoint(65, 10000001, ""    , (' + startPointX + ')*FCN, (' + startPointY + ')*FCN, (0)*FCN)');
                        for (k = 1; k < entity.vertices.length; k++) {
                            if (entity.vertices[k - 1]['bulge']) {    // Add curve
                                mlid = outputProgramm.length + 10000000;
                                curveBulge = +entity.vertices[k - 1]['bulge'];
                                curveDirection = Math.sign(curveBulge) == -1 ? 1 : 2;
                                startPointX = +entity.vertices[k - 1]['x'];
                                startPointY = +entity.vertices[k - 1]['y'];
                                endPointX = +entity.vertices[k]['x'];
                                endPointY = +entity.vertices[k]['y'];

                                if (curveDirection == 1) {
                                    if (+endPointX > +startPointX) {
                                        curveSolution = 1
                                    }
                                    if (+endPointX < +startPointX) {
                                        curveSolution = 0
                                    }
                                    if (+endPointX == +startPointX) {
                                        if (+endPointY > +startPointY) {
                                            curveSolution = 1
                                        }
                                        if (+endPointY <= +startPointY) {
                                            curveSolution = 0
                                        }
                                    }
                                }
                                if (curveDirection == 2) {
                                    if (+endPointX > +startPointX) {
                                        curveSolution = 0
                                    }
                                    if (+endPointX < +startPointX) {
                                        curveSolution = 1
                                    }
                                    if (+endPointX == +startPointX) {
                                        if (+endPointY > +startPointY) {
                                            curveSolution = 0
                                        }
                                        if (+endPointY <= +startPointY) {
                                            curveSolution = 1
                                        }
                                    }
                                }
                                curveRadius = Math.abs((Math.sqrt((endPointX - startPointX) * (endPointX - startPointX) + (endPointY - startPointY) * (endPointY - startPointY)) / 2 ) / Math.sin(4 * Math.atan(curveBulge) / 2)).toFixed(2);
                                outputProgramm.push('   @ ARC_EPRA, "", "", ' + mlid + ' : ' + endPointX + ', ' + endPointY + ', ' + curveRadius + ', ' + curveDirection + ', 0, 0, 0, 0, 0, ' + curveSolution);
                                outputVBScript.push('Call ProgBuilder.AddArcEPRA(80, ' + mlid + ', ""    , (' + endPointX + ')*FCN, (' + endPointY + ')*FCN, (' + curveRadius + ')*FCN, ' + curveDirection + ', (0)*FCN, (0)*FCN, scOFF, (0)*FCN, 0, ' + curveSolution + ')');
                            } else {
                                mlid = outputProgramm.length + 10000000;
                                endPointX = +entity.vertices[k]['x'];
                                endPointY = +entity.vertices[k]['y'];
                                outputProgramm.push('   @ LINE_EP, "", "", ' + mlid + ' : ' + endPointX + ', ' + endPointY + ', 0, 0, 0, 0, 0, 0, 0');
                                outputVBScript.push('Call ProgBuilder.AddLineEP(66, ' + mlid + ', ""    , (' + endPointX + ')*FCN, (' + endPointY + ')*FCN, (0)*FCN, (0)*FCN, scOFF, (0)*FCN, 0, 0, NO)');
                            }
                        }
                        mlid = outputProgramm.length + 10000000;
                        outputProgramm.push('   @ ENDPATH, "", "", ' + mlid + ' :');
                        outputVBScript.push('Call ProgBuilder.EndPath(64, ' + mlid + ', ""    ): SIDE = -1');
                    }
                }
            }
        }

        for (propName in tdcodes) {
            if (tdcodes.hasOwnProperty(propName)) {
                outputTdcodes.push(tdcodes[propName])
            }
        }
        outputVBScript.push('End Sub');

        console.log(outputFile);
        fs.appendFileSync(outputFile, outputHeader.join('\n'), {encoding: 'utf8', mode: '0644'});
        fs.appendFileSync(outputFile, outputVariables.join('\n'), {encoding: 'utf8', mode: '0644'});
        fs.appendFileSync(outputFile, outputProgramm.join('\n'), {encoding: 'utf8', mode: '0644'});
        fs.appendFileSync(outputFile, outputVBScriptHeader.join('\n'), {encoding: 'utf8', mode: '0644'});
        fs.appendFileSync(outputFile, outputVBScript.join('\n'), {encoding: 'utf8', mode: '0644'});
        fs.appendFileSync(outputFile, outputTdcodes.join('\n'), {encoding: 'utf8', mode: '0644'});
    });
}
function outputHolzma() {

    var holzmaArr,
        borderList,
        borderObj = {},
        materialList,
        materialObj = {},
        holzmaObj = [],
        holzmaPan,
        startCommentIndex,
        endCommentIndex,
        i;

    var outputFile = path.join(O_PATH, process.argv[2]) + '.csv';

    try {
        holzmaArr = fs.readFileSync(path.join(I_PATH, process.argv[2]) + '.dkm')
            .toString()
            .split('\r\n');
    } catch (err) {
        if (err.code == 'ENOENT') {
            console.log(path.join(I_PATH, process.argv[2]) + '.dkm' + ' - no such file or directory');
            return;
        }
    }

    borderList = fs.readFileSync('C:/Program Files/Elecran/3D-Constructor 3/Inbase/borders.ini')
        .toString()
        .split('\r\n');

    for (i = 0; i < borderList.length; i++) {
        if (borderList[i].indexOf('=') != -1) {
            borderObj[borderList[i].substring(0, borderList[i].indexOf('='))] = borderList[i].substring(borderList[i].indexOf('=') + 1);
        }
    }

    materialList = fs.readFileSync('C:/Program Files/Elecran/3D-Constructor 3/Inbase/listmat.ini')
        .toString()
        .split('\r\n');

    for (i = 0; i < materialList.length; i++) {
        if (materialList[i].indexOf('=') != -1) {
            materialObj[materialList[i].substring(0, materialList[i].indexOf('='))] = materialList[i].substring(materialList[i].indexOf('=') + 1);
        }
    }

    for (i = 0; i < holzmaArr.length; i++) {

        if (holzmaArr[i].indexOf('[Detal') == 0) {
            holzmaObj.push({});
            holzmaPan = holzmaObj[holzmaObj.length - 1];
        }
        if (holzmaArr[i].indexOf('Shifr=') == 0) {
            startCommentIndex = holzmaArr[i].indexOf('{');
            endCommentIndex = holzmaArr[i].indexOf('}');
            if (startCommentIndex != -1 && endCommentIndex != -1) {
                holzmaPan.comment = '(smotri cherteg)';
                holzmaPan.shifr = holzmaArr[i].substring(endCommentIndex + 1);
            } else {
                holzmaPan.comment = '';
                holzmaPan.shifr = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
            }
        }
        if (holzmaArr[i].indexOf('W=') == 0) {
            holzmaPan.w = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
        }
        if (holzmaArr[i].indexOf('H=') == 0) {
            holzmaPan.h = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
        }
        if (holzmaArr[i].indexOf('Quantity=') == 0) {
            holzmaPan.quantity = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
        }
        if (holzmaArr[i].indexOf('Material=') == 0) {
            holzmaPan.material = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
        }
        if (holzmaArr[i].indexOf('Rotate=') == 0) {
            holzmaPan.rotate = holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1);
        }
        if (holzmaArr[i].indexOf('RB=') == 0) {
            holzmaPan.rb = borderObj[holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1)] + holzmaPan.comment;
        }
        if (holzmaArr[i].indexOf('BB=') == 0) {
            holzmaPan.bb = borderObj[holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1)] + holzmaPan.comment;
        }
        if (holzmaArr[i].indexOf('LB=') == 0) {
            holzmaPan.lb = borderObj[holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1)] + holzmaPan.comment;
        }
        if (holzmaArr[i].indexOf('TB=') == 0) {
            holzmaPan.tb = borderObj[holzmaArr[i].substring(holzmaArr[i].indexOf('=') + 1)] + holzmaPan.comment;
        }
    }
    fs.writeFileSync(OUTPUT_FILE_HOLZMA, JSON.stringify(holzmaObj, null, 3));

    holzmaObj.forEach(function (item, i) {
        fs.appendFileSync(outputFile, ';' + item.shifr + ';' + i + ';' + ';' + item.quantity + ';' + item.w + ';' + item.h + ';' + materialObj[item.material] + ';' + '1;0;' + (item.tb || '') + ';' + (item.bb || '') + ';' + (item.lb || '') + ';' + (item.rb || '') + '\n', {
            encoding: 'utf8',
            mode: '0644'
        });
    });
    console.log(outputFile);
}