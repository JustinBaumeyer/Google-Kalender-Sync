const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,shared,alice,bob}=require('./helpers/apps-script.cjs');

function fixture({extra='',start='20990101T080000Z',end='20990101T090000Z',original={dateTime:'2099-01-02T08:00:00Z'}}={}) {
  const app=shared(),c=load(app,alice),patches=[];
  const settings={sourceCalendars:[['https://example.com/feed.ics','Dienstplan',5,'calendar-'+alice]]};
  c.saveAppSettings({settings});
  const feed='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:series\r\nDTSTART'+(start.includes('VALUE=DATE')?';':':')+start+'\r\nDTEND'+(end.includes('VALUE=DATE')?';':':')+end+'\r\nRRULE:FREQ=DAILY;COUNT=3\r\nSUMMARY:Serie\r\n'+extra+'END:VEVENT\r\nEND:VCALENDAR';
  c.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=>feed})};
  let parent,cancelled;
  c.Calendar.Events={
    list:()=>({items:parent?[parent,...(cancelled?[cancelled]:[])]:[]}),
    insert:event=>{parent=JSON.parse(JSON.stringify({...event,id:'google-series'}));return parent;},
    update:()=>assert.fail('Unchanged series must not be rewritten'),
    patch:(event,calendar,id)=>{patches.push({event,calendar,id});cancelled=null;return event;},
    remove:()=>assert.fail('Existing series must not be deleted'),
  };
  c.startSync();
  cancelled={id:'cancelled-instance',status:'cancelled',recurringEventId:'google-series',originalStartTime:original};
  return {app,c,patches,parent,cancelled,settings,setParentOwner:email=>{parent.extendedProperties.private.syncOwner=email;}};
}

test('a deleted occurrence is restored from an unchanged subscription series and counted once',()=>{
  const {c,patches}=fixture();
  const state=c.startSync();
  assert.equal(patches.length,1);assert.equal(state.lastStats.added,1);
  assert.equal(patches[0].id,'cancelled-instance');assert.equal(patches[0].event.status,'confirmed');
  assert.equal(patches[0].event.start.dateTime,'2099-01-02T08:00:00Z');
  assert.equal(patches[0].event.recurringEventId,undefined);
  assert.equal(c.startSync().lastStats.added,0);assert.equal(patches.length,1);
});

test('source exclusions and dates outside the recurrence are never restored',()=>{
  for(const options of [
    {extra:'EXDATE:20990102T080000Z\r\n'},
    {original:{dateTime:'2099-01-05T08:00:00Z'}},
  ]) {
    const {c,patches}=fixture(options);c.startSync();assert.equal(patches.length,0);
  }
});

test('restoration leaves a source-cancelled occurrence alone',()=>{
  const {c,parent,cancelled,patches}=fixture();
  const text='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:series\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T090000Z\r\nRRULE:FREQ=DAILY;COUNT=3\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nUID:series\r\nRECURRENCE-ID:20990102T080000Z\r\nSTATUS:CANCELLED\r\nEND:VEVENT\r\nEND:VCALENDAR';
  c.calendarEvents=[parent];c.cancelledCalendarEvents=[cancelled];
  c.restoreDeletedOccurrences_(new c.ICAL.Component(c.ICAL.parse(text)).getAllSubcomponents('vevent'),'Europe/Berlin');
  assert.equal(patches.length,0);
});

test('restoration uses a moved source exception at its new time',()=>{
  const {c,parent,cancelled,patches}=fixture();
  const text='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:series\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T090000Z\r\nRRULE:FREQ=DAILY;COUNT=3\r\nSUMMARY:Serie\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nUID:series\r\nRECURRENCE-ID:20990102T080000Z\r\nDTSTART:20990102T100000Z\r\nDTEND:20990102T110000Z\r\nSUMMARY:Verschoben\r\nEND:VEVENT\r\nEND:VCALENDAR';
  c.calendarEvents=[parent];c.cancelledCalendarEvents=[cancelled];
  c.restoreDeletedOccurrences_(new c.ICAL.Component(c.ICAL.parse(text)).getAllSubcomponents('vevent'),'Europe/Berlin');
  assert.equal(patches.length,1);assert.equal(patches[0].event.start.dateTime,'2099-01-02T10:00:00Z');
  assert.equal(patches[0].event.summary,'Verschoben');
});

test('restoration respects creation settings, future-only filtering and account ownership',()=>{
  for(const mode of ['creation-disabled','other-owner','past']) {
    const f=fixture(mode==='past'?{start:'20200101T080000Z',end:'20200101T090000Z',original:{dateTime:'2020-01-02T08:00:00Z'}}:{});
    if(mode==='other-owner')f.setParentOwner(bob);
    if(mode!=='other-owner')f.c.saveAppSettings({settings:{...f.settings,...(mode==='past'?{onlyFutureEvents:true}:{addEventsToCalendar:false})}});
    f.c.startSync();assert.equal(f.patches.length,0);
  }
});

test('deleted all-day and floating-time occurrences use the source and calendar timezone',()=>{
  const allDay=fixture({start:'VALUE=DATE:20990101',end:'VALUE=DATE:20990102',original:{date:'2099-01-02'}});
  allDay.c.startSync();assert.equal(allDay.patches[0].event.start.date,'2099-01-02');
  const floating=fixture({start:'20990101T080000',end:'20990101T090000',original:{dateTime:'2099-01-02T07:00:00Z'}});
  floating.c.startSync();assert.equal(floating.patches[0].event.start.dateTime,'2099-01-02T08:00:00');
  assert.equal(floating.patches[0].event.start.timeZone,'Europe/Berlin');
});
