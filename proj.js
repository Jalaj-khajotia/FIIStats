console.log('Application loaded');

var XLSX = require('xlsx');
var fileName = '0430.csv';
console.log('File used is ' + fileName);
var workbook = XLSX.readFile(fileName);
var sheet_name_list = workbook.SheetNames;

sheet_name_list.forEach(function (y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var excel = [];
    for (z in worksheet) {
        if (z[0] === '!')
            continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0, tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if (row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if (!excel[row])
            excel[row] = {};
        excel[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    excel.shift();
    excel.shift();

    mydata = [];
    var winners = [];
    var i = 0,
    limit = 0;

    const readline = require('readline').createInterface({
        input: process.stdin,
        output: process.stdout
    });
	console.log('');
    console.log('Participant wise Open Interest (no. of contracts) in Equity Derivatives are');
    showFIIData();

    function showFIIData() {
        var fiiData = excel[2],
        futIndexLong,
        futIndexShort,
        futStockShort,
        futStockLong,
        OptionIndexCallLong,
        OptionIndexCallShort,
        OptionIndexPutLong,
        OptionIndexPutShort,
        OptionStockCallLong,
        OptionStockCallShort,
        OptionStockPutLong,
        OptionStockPutShort;
        if (fiiData['Client Type'] == 'FII') {
            futIndexLong = fiiData['Future Index Long'];
            futIndexShort = fiiData['Future Index Short'];
            futStockShort = fiiData['Future Stock Short\t'];
            futStockLong = fiiData['Future Stock Long'];
            OptionIndexCallLong = fiiData['Option Index Call Long'];
            OptionIndexCallShort = fiiData['Option Index Call Short'];
            OptionIndexPutLong = fiiData['Option Index Put Long'];
            OptionIndexPutShort = fiiData['Option Index Put Short'];
            OptionStockCallLong = fiiData['Option Stock Call Long'];
            OptionStockCallShort = fiiData['Option Stock Call Short'];
            OptionStockPutLong = fiiData['Option Stock Put Long'];
            OptionStockPutShort = fiiData['Option Stock Put Short'];

            var futIndexLongPer = futIndexLong / (futIndexLong + futIndexShort);
            var futStockShortPer = futStockShort / (futStockShort + futStockLong);
            var OptionIndexCallLongPer = OptionIndexCallLong / (OptionIndexCallLong + OptionIndexCallShort);
            var OptionIndexPutLongPer = OptionIndexPutLong / (OptionIndexPutLong + OptionIndexPutShort);
            var OptionStockCallLongPer = OptionStockCallLong / (OptionStockCallLong + OptionStockCallShort);
            var OptionStockPutLongPer = OptionStockPutLong / (OptionStockPutLong + OptionStockPutShort);

            console.log('FII Future Index Long is ' + Math.round(futIndexLongPer * 100) + '%');
            console.log('FII Option Index Call Long is ' + Math.round(OptionIndexCallLongPer * 100) + '%');
            console.log('FII Option Index Put Long is ' + Math.round(OptionIndexPutLongPer * 100) + '%');
            console.log('FII Future Stock Long is ' + (100 - Math.round(futStockShortPer * 100)) + '%');
            console.log('FII Option Stock Call Long is ' + Math.round(OptionStockCallLongPer * 100) + '%');
            console.log('FII Option Stock Put Long is ' + Math.round(OptionStockPutLongPer * 100) + '%');
            readline.close();
        }
    }

    function GapUpGainers() {
        gapUpList = [];
        var i = 0;
        readline.question(`Enter the Gap Up %? `, (gapUp) => {
            readline.question('Enter Gainer minimum %, Default:0% ', (gain) => {
                excel.forEach(function (cell) {
                    var gapupPercentage = (cell.OPEN - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    var riseFall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    if (gapupPercentage > 0 && gapupPercentage > gapUp && riseFall >= gain && cell.TOTTRDQTY > 10000) {

                        var percentage = Math.round(riseFall * 100) / 100;
                        var valuegapup = Math.round(gapupPercentage * 100) / 100;
                        gapUpList[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: percentage,
                            GapUp: valuegapup
                        };
                    }

                });
                console.log('');
                console.log('Listing stocks with gap up > ' + gapUp + '% & gained > ' + gain + '%');
                console.log('');
                console.log('  ' + 'Stock Name' + '\t' + 'GapUp %' + '\t' + 'Increase');

                gapUpList.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t' + stock.GapUp + '\t' + stock.Percentage);
                })
                readline.close();
            })
        })
    }

    function DailyGainers() {
        readline.question(`Enter the lower limit? `, (lower) => {
            readline.question('Enter Upper Limit, Default:100% ', (upper) => {
                excel.forEach(function (cell) {
                    var riseFall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    var gapUp = (cell.OPEN - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    var roundGapUp = Math.round(gapUp * 100) / 100;

                    mydata[cell.SYMBOL] = Math.round(riseFall * 100) / 100;
                    limit = upper == 0 ? 100 : upper;
                    if (mydata[cell.SYMBOL] >= lower && mydata[cell.SYMBOL] < upper && cell.TOTTRDQTY > 10000) {
                        // winners[cell.SYMBOL] = mydata[cell.SYMBOL];

                        winners[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: mydata[cell.SYMBOL]
                        }
                    };
                });
                winners.sort(function (a, b) {
                    return a.Percentage - b.Percentage;
                });

                console.log('');
                console.log('Listing stocks which rose > ' + lower + '% but are lower than  < ' + upper + '%');
                console.log('');
                console.log('  ' + 'Stock Name' + '\t' + 'Increase');
                winners.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t' + stock.Percentage);
                })
                // console.log(winners);
                readline.close()

            });

        });
    }

});