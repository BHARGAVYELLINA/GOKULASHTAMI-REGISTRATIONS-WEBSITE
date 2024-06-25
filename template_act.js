var memberContainer = document.getElementById("memberContainer");
var memberTemplate = document.getElementById("member");
//var members = [];
var count = 1;
var minregs=3;
var maxregs=6;
var objarr =[];

function addMember()
{
  if (count<=maxregs){
      var newMember = memberTemplate.content.cloneNode(true);
      head = newMember.getElementById("head");
      head.innerHTML = head.innerHTML + count.toString();
      count++;
      //newMember.id='m'+count.toString();
      memberContainer.appendChild(newMember);
      //members.push(newMember);
  }
  else{
    alert("Maximum 6 members allowed");
  }
}

function mincheck(){
  if(count<=minregs){
    alert("Minimum 3 members required");
    event.preventDefault();
  }
}

document.addEventListener('DOMContentLoaded', function () {
    addMember();
});

function sub() {
    var members = document.querySelectorAll(".form");
    var formDataArray = [];

    for (let i = 0; i < members.length; i++) {
        var formData = new FormData(members[i]);
        var formDataObj = {};

        formData.forEach((value, key) => {
            formDataObj[key] = value;
        });

        formDataArray.push(formDataObj);
    }

    $.post(actionLink, JSON.stringify(formDataArray), (data, status) => {
        console.log(data);
    });
}