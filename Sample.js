//Acquiring all the required modules
const puppeteer = require("puppeteer");
let fs = require("fs");
let path = require("path");
let xlsx=require("xlsx");
const { data } = require("cheerio/lib/api/attributes");


let page;
(async function fn() {
    // Initialisiing and launching Puppeteer browser
    let browser = await puppeteer.launch({
        headless: false, defaultViewport: null,
        args: ["--start-maximized"],
    })
    // Google Page Navigation
    page = await browser.newPage();
    await page.goto("https://www.google.com");
    console.log("Visited Google HomePage!");

    //Searching Bookmyshow on Google
    await page.type("input[title='Search']","bookmyshow");
    await page.keyboard.press('Enter');
    console.log("Bookmyshow searched on Google");

    //Clicking on the first search result.
    await page.waitForSelector(".LC20lb.DKV0Md",{visible:true});
    await page.click(".LC20lb.DKV0Md",{visible:true});
    console.log("Opened Bookmyshow.com");

    //Click on cancel personalized notifications
    await page.waitForSelector('#wzrk_wrapper button[id="wzrk-cancel"]', {visible:true});
    await page.click('#wzrk_wrapper button[id="wzrk-cancel"]', {visible:true},{delay : 1200});
    console.log("Clicked on Not Now");

    //To click on city(DELHI-NCR)
    let cityList = await page.$$('div[id="modal-root"] ul li');
    console.log("Number of cities we can choose:", cityList.length);

    await cityList[1].click({delay:1200})
    console.log("Clicked on Delhi");

    //To see list of all recommended movies currently 
    await page.waitForSelector('div[id="super-wrapper"] div[class="style__StyledText-sc-7o7nez-0 cZrjnA"]');
    await page.click('div[id="super-wrapper"] div[class="style__StyledText-sc-7o7nez-0 cZrjnA"]');
    console.log("CLicked on See All to get the full recommended list");
    
    await autoScroll(page);
    //Number of recommended movies on the Web Page
    let movieNameList = await page.$$('div[id="super-wrapper"] div[class="style__StyledText-sc-7o7nez-0 gBgfCW"]',);
    console.log("Number of Movies in the recommended list:",movieNameList.length);
    
    // Obtaining links gor all the movies in the Movie List present on the Web Page
    let linkarr = await page.$$eval('[class="commonStyles__LinkWrapper-sc-133848s-11 style__CardContainer-sc-1ljcxl3-1 Xdzak"]',assetLinks => assetLinks.map(link => link.href));

    let linkarr2=[];
    for(let i=0;i<linkarr.length-1;i++){
        linkarr2[i] = linkarr[i+1];
    }
    //Object that contains all the elements
    let Movieobjfinal=[]

    //Iterating on the movie link array to visit each movie detail page to obtain data 
    for(let i = 0;i<linkarr2.length;i++){
        await page.goto(linkarr2[i]);
        let movietitle = await page.$$eval('h1[class = "styles__EventHeading-sc-qswwm9-6 dEOMtn"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let contentarr1 = await page.$$eval('[class="styles__LinkedElementsContainer-sc-2k6tnd-4 ehHDzd"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let contentarr2 = await page.$$eval('[class = "styles__EventAttributesContainer-sc-2k6tnd-1 hSMSQi"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let RatingbyPeople = await page.$$eval('[class="styles__Title-sc-ec6ph5-3 eADFBc"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let cast = await page.$$eval('section[id="component-3"] [class="styles__Title-sc-17p4id8-1 jfGPxs"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let GetReviewTags = await page.$$eval('[class ="styles__PillsComponent-sc-h9dyha-0 ffhovO"]',anchors => { return anchors.map(anchor => anchor.textContent)});
        let movieTitle =movietitle[0];
        let RatingByPeople = RatingbyPeople[0];
        let formats = contentarr1[0];
        let languages = contentarr1[1];
        let Cast = "";
        for(let i =0;i<cast.length;i++){
            Cast = Cast+cast[i]+", ";
        }
        let ReviewTags = ""
        for(let i =0;i<GetReviewTags.length;i++){
            ReviewTags = ReviewTags+GetReviewTags[i]+", ";
        }
        contentarr2 = contentarr2[1];
        contentarr2 = contentarr2.split(" • ");
        let movieTime = contentarr2[0];
        let movieGenres = contentarr2[1];
        let movieRating = contentarr2[2];
        let movieReleaseDate = contentarr2[3];
        let MovieObj = {
            movieTitle,
            movieTime,
            formats,
            languages,
            movieGenres,
            movieRating,
            movieReleaseDate,
            RatingByPeople,
            Cast,
            ReviewTags
        }
        // Pushing data for each movie into an object
        Movieobjfinal.push(MovieObj);

    }
    excelWriter("Bookmyshow.xlsx",Movieobjfinal,"sheet-1");


})();

async function autoScroll(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;

                if(totalHeight >= scrollHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
}

function excelWriter(filePath, json, sheetName){
    let newWB=xlsx.utils.book_new();
    let newWS=xlsx.utils.json_to_sheet(json);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    xlsx.writeFile(newWB,filePath);
}



