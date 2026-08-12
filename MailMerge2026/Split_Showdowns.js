// This script is to be used in Adobe Acrobat Pro DC
// It inserts pages to prepare showdowns for printing
// The script searches for the word "PAGE" followed by a number and then the same number
// If the numbers are odd, the page is considered a showdown page
// The script inserts a blank page before each showdown page
// The script then saves the new file with "_split" appended to the original file name
var splitShowdown = app.trustedFunction(function() {
app.beginPriv();
var oddPages = [];
var newFileName = this.path.replace(/\.pdf$/, "_split.pdf");
var nPages = this.numPages;
for (var i = 0; i < nPages; i++) {
for (var n = 0; n < this.getPageNumWords(i); n++) {
var currentWord = this.getPageNthWord(i,n);
if (currentWord=="PAGE" && this.getPageNthWord(i,n+1).length==1) {
var advOne = Number(this.getPageNthWord(i, n+1));
var advTwo = this.getPageNthWord(i, n+2);
var advThree = Number(this.getPageNthWord(i, n+3));
if (advOne == advThree && advThree%2!=0) {
// console.println(currentWord + " " + advOne + " " + advTwo + " " + advThree);
// console.println("odd page found: " + i);
oddPages.push(i);
}
break;
}
}
}
var newDoc = app.newDoc();
var newDocPath = newDoc.path;
for (var i = oddPages.length - 1; i >= 0; i--) {
this.insertPages({
nPage: oddPages[i],
cPath: newDocPath,
nStart: 0
});
}
newDoc.closeDoc(true);
this.saveAs(newFileName);
console.println("new file saved as " + newFileName);
app.endPriv();
});