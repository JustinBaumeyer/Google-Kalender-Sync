// Timing fields transcribed from the supplied ROSTER legend; vehicle labels anonymized.
const clock=(hourOfDay,minute=0)=>({hourOfDay,minute});
const definitions={
  K9T:{shortName:'K9T',longName:'Station 1 KTW 1 Tag',from:clock(7),to:clock(19),followingDaysTo:0,shiftTypeName:'Tagdienst',durationRawInMinutes:720,durationInMinutes:659,isFlexiTime:false},
  K9N:{shortName:'K9N',longName:'Station 1 KTW 1 Nacht',from:clock(19),to:clock(7),followingDaysTo:1,shiftTypeName:'Nachtdienst',durationRawInMinutes:720,durationInMinutes:659,isFlexiTime:false},
  ITW:{shortName:'ITW',longName:'ITW werktags Präsenz',from:clock(7,30),to:clock(16),followingDaysTo:0,shiftTypeName:'Frühdienst',durationRawInMinutes:510,durationInMinutes:510,isFlexiTime:false},
  B24:{shortName:'B24',longName:'Rufbereitschaft 24',from:clock(6),to:clock(6),followingDaysTo:1,shiftTypeName:'Rufbereitschaft',durationRawInMinutes:1440,durationInMinutes:0,isFlexiTime:false},
  AHD:{shortName:'AHD',longName:'Lehrgang',from:clock(7,30),to:clock(17,30),followingDaysTo:0,shiftTypeName:'Ausbildung',durationRawInMinutes:600,durationInMinutes:555,elements:[
    {from:clock(7,30),to:clock(17,30),followingDaysTo:0,shiftTypeName:'Regeldienst'},
    {from:clock(12),to:clock(12,45),followingDaysTo:0,shiftTypeName:'Pause'}
  ],isFlexiTime:false}
};
const definition=(code,overrides={})=>Object.assign(JSON.parse(JSON.stringify(definitions[code])),overrides);
module.exports={clock,definition};
