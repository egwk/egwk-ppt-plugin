/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var locales = {
  en: {
    welcome: 'Find paragraphs in the writings of Ellen White.',
    find: 'Find',
    searchBoxHint: 'Find...',
    searchInProgress: 'Looking up phrase...',
    noResult: 'No matches found',
    insert: 'Insert',
    errorTryLater: 'Error during search. <br>Try again later!'
  },
  hu: {
    welcome: 'Keress az Ellen Gould White Könyvtár bekezdéseiben!',
    find: 'Keresés',
    searchBoxHint: 'Keresés...',
    searchInProgress: 'Keresés...',
    noResult: 'A megadott kifejezésre nincsen találat',
    insert: 'Beszúrás',
    errorTryLater: 'Hiba a keresés során. <br>Próbáld újra később!'
  }
}
var ui = locales.en;
var running = false;
var text = "";
// var startTime; // futási idő méréshez 1/3

Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    var uiLang = Office.context.displayLanguage;
    switch (uiLang) {
      case 'hu-HU':
        ui = locales.hu;
        break;
      default:
        ui = locales.en;
        break;
    }
    var appHtml = `
    <div class="top-search">
      <div class="searchbox-container">
        <form onsubmit="return false">
          <div class="BorderGray">
            <input class="DefaultText" type="text" id="searchBox" placeholder="${ui.searchBoxHint}">
            <input title="${ui.find}" id="submitSearch" class="BtnSearchButton" type="image" alt="${ui.find}" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAASCAMAAABhEH5lAAAAbFBMVEW9vb3AwMDb29vo6OjIyMjk5OTKysrBwcHu7u709PTOzs7R0dHS0tLV1dXZ2dnDw8Pd3d3e3t7g4ODh4eG/v7++vr7p6enq6urs7Ozt7e3MzMzw8PDx8fHNzc3+/v7////GxsbQ0NDPz8/i4uIwwMXbAAAAaUlEQVR4Xp2QRw6EQAwEx/YEcs6Z3f//kXAxMCfoS0klS25bzFbeKSeQK90UQkQo/YtqwdkxjRkrH0+QZKXVCTe2pirJqimHA17K6i+KvFG/jjciQNAtOuFeYQHq3t41pn4eRP2nT1jZABlDIuI2GxSeAAAAAElFTkSuQmCC">
          </div>
        </form>
      </div>
    </div>
    <main id="app-body" class="ms-welcome__main">
      <div id="welcome-text" class="center">
        <p>
          ${ui.welcome}
        </p>
        <p>
          <a href="http://egw.hu">http://egw.hu</a>
        </p>
      </div>
      <div id="search-spinner" style="display: none;">
        <div>
          ${ui.searchInProgress}
          <div class="spinner"></div>
        </div>
      </div>
      <ul class="ms-List"  id="search-results">
      </ul>
    </main>
    `;
    document.getElementById('app-container').innerHTML = appHtml;

    var timer;
    document.getElementById("searchBox").addEventListener('keyup', function(e){ 
      clearTimeout(timer);
      timer = setTimeout(function(){
        // startTime = window.performance.now(); // futási idő mérés 2/4
        prepareSearch();
      }, 1000);
    });
    document.getElementById("submitSearch").onclick = function(){
      running = false;
      // startTime = window.performance.now(); // futási idő mérés 3/4
      prepareSearch();
    }
  }
});

async function prepareSearch(){
  var term = document.getElementById("searchBox").value;

  // az instant search (on keyup) ne keressen, ha már nyomtunk entert
  if (running === term) {
    return false;
  }

  running = term;

  // welcome text és előző keresési eredmények törlése
  const welcomeText = document.querySelector('#welcome-text');
  welcomeText.style.display = "none";
  
  const myNode = document.getElementById("search-results");
  while (myNode.firstChild) {
    myNode.removeChild(myNode.firstChild);
  }
  showSpinner();
  var searchResults = [];
  var translationResults = await search(term);
  async function buildTranslation(){
    translationResults.forEach( async item => {
      item.refcode_short = await getRefCode(item.para_id);
    });
    return;
  }
  await buildTranslation();

  // keresés az eredeti nyelvű gyűjteményben + refcode kezelés
  let originalResults = await search(term, false);
  // ha refcodera történik a keresés, a fordításokat is adja vissza
  let refcodeResults = [];
  await asyncForEach(originalResults, async (item, index) => {
    if (item.refcode_short.toLowerCase() == term.toLowerCase()) {
      refcodeResults = await findByRefcode(term);
      refcodeResults.forEach(translatedResult => {
        translatedResult.refcode_short = item.refcode_short;
      });
    };
  });

  searchResults = translationResults.concat(refcodeResults, originalResults);
  listResults(searchResults);
  showSpinner(false);
}

async function asyncForEach(array, callback) {
  for (let index = 0; index < array.length; index++) {
    await callback(array[index], index, array);
  }
}

function showSpinner(state = true) {
  document.querySelector("#search-spinner").style.display = state ? "block" : "none";
}

async function search(term, translation = true){
    // Set up our HTTP request
    var xhr = new XMLHttpRequest();

    return new Promise((resolve, reject) => {
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status >= 200 && xhr.status < 300) {
          text = JSON.parse(xhr.responseText).data;
          resolve(text);
        } else {
          printError(ui.errorTryLater);
          console.log('Hiba történt az xhr során. (search)');
        }
      };
      xhr.onerror = function() {
        console.log(xhr.statusText);
      }
      var urlPart = translation ? '/translation' : '';
      var url = 'https://api.white-konyvtar.hu/reader/search' + urlPart + '?query=' + encodeURIComponent(term);
      xhr.open('GET', url);
      xhr.send();
    });
}

function getRefCode(paraId){
  var xhr = new XMLHttpRequest();
  return new Promise((resolve, reject) => {
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      if (xhr.status >= 200 && xhr.status < 300) {
        var result = JSON.parse(xhr.responseText).data;
        var refCode = result[0].refcode_short;
        resolve(refCode);
      } else {
        console.log('Hiba történt az xhr során. (getRefCode)', xhr.status, xhr.statusText);
      }
    };
    xhr.onerror = function() {
      console.log(xhr.statusText);
    }
    var url = 'https://api.white-konyvtar.hu/reader/paragraph/' + encodeURIComponent(paraId);
    xhr.open('GET', url);
    xhr.send();
  });
}

function findByRefcode(term){
  var xhr = new XMLHttpRequest();
  return new Promise((resolve, reject) => {
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      if (xhr.status >= 200 && xhr.status < 300) {
        var result = JSON.parse(xhr.responseText).data;
        var translations = result[0].translations;
        resolve(translations);
      } else {
        console.log('Hiba történt az xhr során. (findByRefcode)');
      }
    };
    xhr.onerror = function() {
      console.log(xhr.statusText);
    }
    var url = 'https://api.white-konyvtar.hu/reader/paragraph/' + encodeURIComponent(term);
    xhr.open('GET', url);
    xhr.send();
  });
}

function listResults(results){
  var resultsCount = Object.keys(results).length;

  // ha van találat
  if (resultsCount > 0) {
    var preparedTemplate = '';
    for (var i=0; i < resultsCount; i++) {
      var template = `
      <li class="ms-ListItem result-li">
        <div class="result-text">
          <p>${results[i].content}</p>
          <div class="fade-effect"></div>
        </div>
        <button class="unhideText">
          <div class="arrowhead down"></div>
        </button>
        <div class="resultButtons">
          <button class="insertText ms-Button">${ui.insert}</button>
          <div class="result-reference">{${results[i].refcode_short}}</div>
        </div>
        <hr style="border: 0; border-top: 1px solid #eaeaea">
      </li>
      `;
      preparedTemplate += template;
    }
    document.getElementById("search-results").innerHTML = preparedTemplate; 
  } else { // ha nincs találat
    console.log("nincs találat");
    var template = `
      <div class="no-result center">${ui.noResult}</div>
    `;
    document.getElementById("search-results").innerHTML += template;
  };

  Array.prototype.forEach.call(document.getElementsByClassName("result-li"), function(el, index, array) {
    // kattintáskor beszúrja a dokumentumba az idézetet
    el.getElementsByClassName("insertText")[0].onclick = function(e){
      insert(results[index].content, results[index].refcode_short);
    };

    // kibontja a szöveget
    el.querySelector(".unhideText").onclick = function(e){
      var resultText = el.querySelector(".result-text");
      if (this.classList.contains('opened')) {
        this.classList.remove('opened');
        resultText.classList.remove('opened');
        // animált összecsukás miatt kell:
        resultText.style.maxHeight = "10rem";
      } else {
        this.classList.add('opened');
        resultText.classList.add('opened');
        // animált kibontás miatt kell:
        var finalHeight = resultText.querySelector("p").offsetHeight;
        resultText.style.maxHeight = finalHeight + "px";
      }
    };
  });

  // csak akkor jelenjen meg ez a fade átmenet, ha van olyan szöveg, ami el van rejtve. 
  // ez itt alul azt ellenőrzi, hogy kilóg-e a 'p' a .result-text-ből.
  const containers = document.querySelectorAll('.result-text');
  Array.prototype.forEach.call(containers, (container) => {  // Loop through each container
    var p = container.querySelector('p');
    var divh = container.clientHeight;
    if (p.offsetHeight > divh) { // Check if the paragraph's height is taller than the container's height. If it is:
      container.classList.add('tooLong');
    }
  });
  // console.log('teljesítve: ', window.performance.now() - startTime); // futási idő mérés 4/4
}

function printError(msg){
  showSpinner(false);
  var template = `<div class="no-result center">${msg}</div>`
  document.getElementById("search-results").innerHTML = template;
}

export async function insert(content, refcode) {
  
  // PowerPoint API beszúrás
  text = content + " " + "{" + refcode + "}";

  Office.context.document.setSelectedDataAsync(text,
    {
      coercionType: Office.CoercionType.Text
    },
    result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error.message);
      }
    }
  );
}

