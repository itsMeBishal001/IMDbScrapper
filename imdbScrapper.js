//download html using axios
//extract information using jsdom
//convert movies to teams
//save movies to excel file using excel4node
//npm init -y
//npm  install minimist
//npm install axios
//npm  install jsdom
//npm install excel4node
//node imdbScrapper.js --url="https://www.imdb.com/search/title/?groups=top_250&sort=user_rating" --excel=imdbBest250.xls

let minimist = require("minimist");
let axios = require("axios");
let jsdom = require("jsdom");
let excel = require("excel4node");
let fs = require("fs");

//taking data from user  using minimist
let args = minimist(process.argv);
//download html using axios

let downloadPromise = axios(args.url);
downloadPromise.then(function (response) {
    let html = response.data;
    let dom = new jsdom.JSDOM(html);
    let doc = dom.window.document;
    let mdc = doc.querySelectorAll("div.lister-item-content");
    let allMovies = [];
    for (let i = 0; i < mdc.length; i++) {
        let movie = {
            slNo: "",
            movieName: "",
            rating: "",
            year: "",
            runTime: "",
            content: "",
            director: "",
            stars: "",
            votes: ""

        };
        //adding sl number
        let slN = mdc[i].querySelector("span.lister-item-index ");
        movie.slNo = slN.textContent;
        //adding movie name
        let mn = mdc[i].querySelector("h3.lister-item-header> a");
        movie.movieName = mn.textContent;
        //adding rating
        let ratin = mdc[i].querySelector("strong")
        movie.rating = ratin.textContent;

        //adding year
        let year = mdc[i].querySelector(".lister-item-year");
        movie.year = year.textContent;
        //adding runtime
        let runtime = mdc[i].querySelector(".runtime")
        movie.runTime = runtime.textContent;
        //adding content
        let contents = [];
        contents = mdc[i].querySelectorAll("p.text-muted");
        let ss = contents[1].textContent;
        let mainC = ss.substring(1);
        movie.content = mainC;
        //adding director and stars

        let dirStars = mdc[i].querySelectorAll("p")
        //let aTAgs=dirStars.querySelectorAll("a");
        // console.log(dirStars[2].querySelectorAll("a").textContent+" "+i);

        //adding director
        movie.director = dirStars[2].querySelector("a").textContent;

        //adding stars
        let stars = dirStars[2].querySelectorAll("a");
        let st = "";
        for (let j = 1; j < stars.length; j++) {
            st += stars[j].textContent + ",";
        }
        //console.log(st+"  "+i);
        movie.stars = st;
        //adding votes
        let votes = mdc[i].querySelectorAll("p.sort-num_votes-visible>span")
        //console.log(votes[1].textContent);
        movie.votes = votes[1].textContent;


        allMovies.push(movie);
    }


    //converting json to stringyfy as we can not write or read json file
    let movieJson = JSON.stringify(allMovies);
    //writeing 
    fs.writeFileSync("allMovies.json", movieJson, "utf-8");

    writeOnExcel(allMovies, args.excel);


}).catch(function (error) {
    console.log(error);
});
function writeOnExcel(allMovies, excelFile) {
    // Create a new instance of a Workbook class
    let wb = new excel.Workbook();
    // Add Worksheets to the workbook
    for (let i = 0; i < allMovies.length; i++) {
        let msheet = wb.addWorksheet(allMovies[i].movieName);
        msheet.cell(1, 1).string("sl No");
        msheet.cell(2, 1).string("movie Name");
        msheet.cell(3, 1).string("rating");
        msheet.cell(4, 1).string("year");
        msheet.cell(5, 1).string("runTime");
        msheet.cell(6, 1).string("content");
        msheet.cell(7, 1).string("director");
        msheet.cell(8, 1).string("stars");
        msheet.cell(9, 1).string("votes");


        msheet.cell(1, 2).string(allMovies[i].slNo);
        msheet.cell(2, 2).string(allMovies[i].movieName);
        msheet.cell(3, 2).string(allMovies[i].rating);
        msheet.cell(4, 2).string(allMovies[i].year);
        msheet.cell(5, 2).string(allMovies[i].runTime);
        msheet.cell(6, 2).string(allMovies[i].content);
        msheet.cell(7, 2).string(allMovies[i].director)
        msheet.cell(8, 2).string(allMovies[i].stars)
        msheet.cell(9, 2).string(allMovies[i].votes)
        let style = wb.createStyle({
            font: {
                color: '#FF0800',
                size: 12,
            }
            
        });
        let bgStyle = wb.createStyle({
            fill: {
              type: 'pattern',
              patternType: 'solid',
              bgColor: '#FFFF00',
              fgColor: '2172d7',
            }
          });

        msheet.cell(1, 1).style(style);
        msheet.cell(2, 1).style(bgStyle);
        msheet.cell(3, 1).style(style);
        msheet.cell(4, 1).style(bgStyle);
        msheet.cell(5, 1).style(style);
        msheet.cell(6, 1).style(bgStyle);
        msheet.cell(7, 1).style(style);
        msheet.cell(8, 1).style(bgStyle);
        msheet.cell(9, 1).style(style);

    }



    wb.write(excelFile);
}
