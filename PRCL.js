async function PRCL() {
    
const query = {
    user: '0503',
    passwd: '', 
    fromDate: '03/23/2024', 
    toDate: '03/23/2024', 
    step: 60, 
    part: '0503', 
    tfa: ''
}

const response = await fetch(`https://webapps.eso.bg/prcl/json/getData.php?${decodeURIComponent(new URLSearchParams(query).toString())}`);
const result = await response.json();
console.log(result);
}

PRCL()
