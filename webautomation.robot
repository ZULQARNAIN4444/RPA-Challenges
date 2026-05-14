*** Settings ***
Library    Autosphere.Browser.Selenium
Library    WebAutomationExcelReader.py

*** Variables ***
${URL}             https://www.theautomationchallenge.com/
${EXCEL_PATH}      B:\\theautomation_chanllenge\\challenge.xlsx

*** Keywords ***
Dismiss Unexpected Alert If Present
    Run Keyword And Ignore Error    Handle Alert    ACCEPT    100ms

Button Exists
    [Arguments]    ${label}
    ${found}    Execute Javascript
    ...    var n=(v)=>(v||'').replace(/\s+/g,' ').trim();
    ...    var v=(e)=>{if(!e)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return s.visibility!=='hidden'&&s.display!=='none'&&r.width>0&&r.height>0;};
    ...    return Array.from(document.querySelectorAll('button')).some(e=>v(e)&&n(e.textContent)==='${label}');
    RETURN    ${found}

Click Button With Events
    [Arguments]    ${label}
    Dismiss Unexpected Alert If Present
    ${clicked}    Execute Javascript
    ...    var n=(v)=>(v||'').replace(/\s+/g,' ').trim();
    ...    var v=(e)=>{if(!e)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return s.visibility!=='hidden'&&s.display!=='none'&&r.width>0&&r.height>0;};
    ...    var b=Array.from(document.querySelectorAll('button')).find(e=>v(e)&&n(e.textContent)==='${label}');
    ...    if(!b)return false;
    ...    b.scrollIntoView({block:'center'});
    ...    var r=b.getBoundingClientRect(),cx=r.left+r.width/2,cy=r.top+r.height/2;
    ...    for(var t of['pointerdown','mousedown','pointerup','mouseup','click']){b.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window,clientX:cx,clientY:cy}));}
    ...    window.__rf_input_map=null;window.__rf_input_map_signature=null;
    ...    return true;
    Should Be True    ${clicked}    Button "${label}" not found/clicked
    Dismiss Unexpected Alert If Present

Fill Field By Label
    [Arguments]    ${label}    ${value}
    ${filled}    Execute Javascript
    ...    var idMap={'Company Name':'company_name','Address':'address','EIN':'ein','Sector':'sector','Automation Tool':'automation_tool','Annual Saving':'annual_saving','Date':'date'};
    ...    var v=(e)=>{if(!e)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return s.visibility!=='hidden'&&s.display!=='none'&&r.width>0&&r.height>0;};
    ...    var prefix=idMap['${label}'];
    ...    if(!prefix)return false;
    ...    var el=Array.from(document.querySelectorAll('input')).find(x=>(x.id||'').startsWith(prefix)&&v(x)&&!x.disabled);
    ...    if(!el)return false;
    ...    el.scrollIntoView({block:'center'});el.focus();el.value='${value}';el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));return true;
    Should Be True    ${filled}    Could not find/fill input for "${label}"

Wait For Next Form
    ${ready}    Execute Javascript
    ...    var v=(e)=>{if(!e)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return s.visibility!=='hidden'&&s.display!=='none'&&r.width>0&&r.height>0;};
    ...    var t=(document.body.innerText||'').replace(/\s+/g,' ').trim();
    ...    if(t.includes('success rate')||t.includes('Certificate of Completion')||t.includes('TRY AGAIN'))return true;
    ...    return Array.from(document.querySelectorAll('input')).filter(e=>v(e)&&!e.disabled).some(e=>!(e.value||'').trim());
    Should Be True    ${ready}    Next form not ready within timeout

Solve Recaptcha If Present
    ${found}    Execute Javascript
    ...    var v=(e)=>{if(!e)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return s.visibility!=='hidden'&&s.display!=='none'&&r.width>0&&r.height>0;};
    ...    return Array.from(document.querySelectorAll('iframe')).some(e=>v(e)&&(e.src||'').includes('recaptcha')&&(e.src||'').includes('anchor'));
    IF    ${found}
        Sleep    1s    Wait for reCAPTCHA iframe to fully load
        Select Frame    xpath://iframe[contains(@src, 'recaptcha') and contains(@src, 'anchor')]
        Sleep    500ms
        Click Element    css:#recaptcha-anchor
        Unselect Frame
        Sleep    2s    Wait for verification to complete
    END

Find 7 Visible Fields
    ${all_inputs}    Get WebElements    xpath://input
    ${visible}    Evaluate    [i for i in $all_inputs if i.is_displayed()]
    ${count}    Evaluate    len($visible)
    Should Be True    ${count} == 7    Expected 7 visible fields, got ${count}
    RETURN    ${visible}

*** Tasks ***
Fill Form With Excel Data
    Open Browser    ${URL}    Chrome
    Set Window Size    1920    1080
    ${config}    Evaluate    json.load(open(r'B:\\theautomation_chanllenge\\config.json', 'r'))    json
    Click Element    xpath://button[contains(text(),'SIGN UP OR LOGIN')]
    Sleep    2s
    Click Element    xpath://button[text()='OR LOGIN']
    Sleep    2s
    Execute Javascript    var e=document.querySelector('input[placeholder="Email"]');e.focus();e.value='${config}[email]';e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))
    Sleep    0.3s
    Execute Javascript    var e=document.querySelector('input[placeholder="Password"]');e.focus();e.value='${config}[password]';e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))
    Sleep    1s
    Click Element    xpath://button[text()='LOG IN']
    Sleep    2s
    Execute Javascript    document.querySelector('div.greyout')?.remove()
    Wait Until Keyword Succeeds    10x    500ms    Click Button With Events    Start
    Sleep    1s
    ${submit_up}    Run Keyword And Return Status    Wait Until Keyword Succeeds    8x    500ms    Button Exists    Submit
    IF    not ${submit_up}
        Wait Until Keyword Succeeds    10x    500ms    Click Button With Events    Start
    END
    Wait Until Keyword Succeeds    10x    500ms    Button Exists    Submit
    ${summary}    Read Excel File    ${EXCEL_PATH}
    ${rows}    Get Excel Data As List
    FOR    ${row}    IN    @{rows}
        Sleep    1s
        Execute Javascript    document.querySelector('div.greyout')?.remove()
        Fill Field By Label    Company Name    ${row}[company_name]
        Fill Field By Label    Address    ${row}[company_address]
        Fill Field By Label    EIN    ${row}[employer_identification_number]
        Fill Field By Label    Sector    ${row}[sector]
        Fill Field By Label    Automation Tool    ${row}[automation_tool]
        Fill Field By Label    Annual Saving    ${row}[annual_automation_saving]
        Fill Field By Label    Date    ${row}[date_of_first_project]
        Execute Javascript    document.querySelector('div.greyout')?.remove()
        Solve Recaptcha If Present
        Wait Until Keyword Succeeds    10x    500ms    Click Button With Events    Submit
        Wait Until Keyword Succeeds    40x    250ms    Wait For Next Form
    END
    Log    ${summary}    console=True
    Sleep    10s
    Close Browser
