function level_1() {
  let x = document.getElementsByClassName('.level2');
  console.log(x)
  if (x.style.display === "none") {
    x.style.display = "block";
  } else {
    x.style.display = "none";
  }
}

function level_2() {
  let x = document.getElementsByClassName('.level3');
  if (x.style.display === "none") {
    x.style.display = "block";
  } else {
    x.style.display = "none";
  }
}