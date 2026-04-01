*** Settings ***
Library           SeleniumLibrary
Library           excel_reader.py
Library           Collections
Library           OperatingSystem

Suite Setup       Open Browser To Challenge
Suite Teardown    Close Browser

*** Variables ***
${URL}            https://www.theautomationchallenge.com/
${BROWSER}        chrome
${EXCEL}          B:/RobotFramework_practice/Scripts/data.xlsx
${CHROME_PROFILE_TEMPLATE_DIR}    B:/RobotFramework_practice/chrome-user-data
${CHROME_RUNTIME_USER_DATA_DIR}    B:/RobotFramework_practice/chrome-user-data-run
${CHROME_PROFILE}    Profile 1

*** Keywords ***
Prepare Chrome Runtime Profile
    Run Keyword And Ignore Error    Remove Directory    ${CHROME_RUNTIME_USER_DATA_DIR}    recursive=True
    Copy Directory    ${CHROME_PROFILE_TEMPLATE_DIR}    ${CHROME_RUNTIME_USER_DATA_DIR}

Open Browser To Challenge
    Prepare Chrome Runtime Profile
    ${options}=    Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys, selenium.webdriver
    ${prefs}=    Create Dictionary    credentials_enable_service=${False}    profile.password_manager_enabled=${False}
    ${arg1}=    Set Variable    --user-data-dir=${CHROME_RUNTIME_USER_DATA_DIR}
    ${arg2}=    Set Variable    --profile-directory=${CHROME_PROFILE}
    ${arg3}=    Set Variable    --disable-notifications
    ${arg4}=    Set Variable    --no-sandbox
    ${arg5}=    Set Variable    --disable-dev-shm-usage

    Call Method    ${options}    add_experimental_option    prefs    ${prefs}
    Call Method    ${options}    add_argument    ${arg1}
    Call Method    ${options}    add_argument    ${arg2}
    Call Method    ${options}    add_argument    ${arg3}
    Call Method    ${options}    add_argument    ${arg4}
    Call Method    ${options}    add_argument    ${arg5}

    Create Webdriver    Chrome    options=${options}
    Go To    ${URL}
    Maximize Browser Window
    Wait Until Element Is Visible    xpath=//button[contains(.,'Start')]    30s

Dismiss Unexpected Alert If Present
    ${status}    ${message}=    Run Keyword And Ignore Error    Handle Alert    ACCEPT    100ms
    IF    '${status}' == 'PASS'
        Log    Alert status: ${status} ${message}
    END

Click Start
    Dismiss Unexpected Alert If Present
    ${start_btn}=    Wait Until Keyword Succeeds    10x    500ms    Get Visible Button    Start
    Execute Javascript    arguments[0].scrollIntoView({block: 'center'});    ARGUMENTS    ${start_btn}
    Sleep    200ms
    Wait Until Keyword Succeeds    4x    500ms    Click Start Element    ${start_btn}
    Execute Javascript    window.__rf_input_map = null; window.__rf_input_map_signature = null;
    Dismiss Unexpected Alert If Present

Activate Challenge
    Click Start
    ${submit_ready}=    Run Keyword And Return Status    Wait Until Keyword Succeeds    8x    500ms    Get Button For Label    Submit
    IF    not ${submit_ready}
        Click Start
    END
    Wait Until Keyword Succeeds    10x    500ms    Get Button For Label    Submit
    Wait Until Keyword Succeeds    10x    500ms    Get Visible Input For Label    Company Name

Wait For Next Form State
    ${script}=    Catenate    SEPARATOR=\n
    ...    const isVisible = (el) => {
    ...        if (!el) return false;
    ...        const style = window.getComputedStyle(el);
    ...        const rect = el.getBoundingClientRect();
    ...        return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
    ...    };
    ...    const completionText = (document.body.innerText || '').replace(/\s+/g, ' ').trim();
    ...    const challengeCompleted = completionText.includes('success rate is') || completionText.includes('Certificate of Completion') || completionText.includes('TRY AGAIN');
    ...    if (challengeCompleted) return true;
    ...    return Array.from(document.querySelectorAll('input'))
    ...        .filter((el) => isVisible(el) && !el.disabled)
    ...        .some((el) => !(el.value || '').trim());
    Wait Until Keyword Succeeds    40x    250ms    Next Form State Should Be Ready    ${script}

Next Form State Should Be Ready
    [Arguments]    ${script}
    ${ready}=    Execute Javascript    ${script}
    Should Be True    ${ready}    Next form state or completion page is not ready yet.

Wait For Completion Page
    ${script}=    Catenate    SEPARATOR=\n
    ...    const text = (document.body.innerText || '').replace(/\s+/g, ' ').trim();
    ...    const hasCertificate = text.includes('Certificate of Completion');
    ...    const hasSuccessRate = text.includes('Your success rate is') || text.includes('success rate is');
    ...    const hasSuccessBanner = text.includes('SUCCESS!');
    ...    return hasCertificate || hasSuccessRate || hasSuccessBanner;
    Wait Until Keyword Succeeds    40x    500ms    Completion Page Should Be Ready    ${script}

Completion Page Should Be Ready
    [Arguments]    ${script}
    ${ready}=    Execute Javascript    ${script}
    Should Be True    ${ready}    Completion page is not ready yet.

Click Submit Button
    Dismiss Unexpected Alert If Present
    ${submit_btn}=    Wait Until Keyword Succeeds    10x    500ms    Get Button For Label    Submit
    Execute Javascript    arguments[0].scrollIntoView({block: 'center'});    ARGUMENTS    ${submit_btn}
    Sleep    200ms
    Wait Until Keyword Succeeds    4x    500ms    Click Visible Element    ${submit_btn}
    Execute Javascript    window.__rf_input_map = null; window.__rf_input_map_signature = null;
    Dismiss Unexpected Alert If Present
    Wait For Next Form State

Get Visible Input For Label
    [Arguments]    ${label}
    ${script}=    Catenate    SEPARATOR=\n
    ...    const label = arguments[0];
    ...    const knownLabels = ['Company Name', 'Address', 'EIN', 'Sector', 'Automation Tool', 'Annual Saving', 'Date'];
    ...    const normalize = (value) => (value || '').replace(/\\s+/g, ' ').trim();
    ...    const isVisible = (el) => {
    ...        if (!el) return false;
    ...        const style = window.getComputedStyle(el);
    ...        const rect = el.getBoundingClientRect();
    ...        return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
    ...    };
    ...    const labelElements = Array.from(document.querySelectorAll('body *'))
    ...        .map((el) => ({ el, text: normalize(el.textContent) }))
    ...        .filter(({ el, text }) => isVisible(el) && knownLabels.includes(text))
    ...        .map(({ el, text }) => ({ el, text, rect: el.getBoundingClientRect(), area: el.getBoundingClientRect().width * el.getBoundingClientRect().height }));
    ...    const labels = knownLabels
    ...        .map((name) => {
    ...            const candidates = labelElements
    ...                .filter((item) => item.text === name)
    ...                .sort((a, b) => (a.area - b.area) || (a.rect.top - b.rect.top) || (a.rect.left - b.rect.left));
    ...            return candidates.length ? { name, element: candidates[0].el, rect: candidates[0].rect } : null;
    ...        })
    ...        .filter(Boolean);
    ...    const inputs = Array.from(document.querySelectorAll('input'))
    ...        .filter((el) => isVisible(el) && !el.disabled && ['text', 'tel', 'search', 'url', 'email', 'number', ''].includes((el.type || '').toLowerCase()));
    ...    if (!labels.length || !inputs.length) return null;
    ...    const inputRects = inputs.map((input) => ({ input, rect: input.getBoundingClientRect() }));
    ...    const signature = JSON.stringify({
    ...        labels: labels.map(({ name, rect }) => [name, Math.round(rect.left), Math.round(rect.top), Math.round(rect.width), Math.round(rect.height)]),
    ...        inputs: inputRects.map(({ rect }) => [Math.round(rect.left), Math.round(rect.top), Math.round(rect.width), Math.round(rect.height)])
    ...    });
    ...    if (window.__rf_input_map && window.__rf_input_map_signature === signature) {
    ...        const cached = window.__rf_input_map[label];
    ...        if (cached && isVisible(cached) && !cached.disabled) return cached;
    ...    }
    ...    const scorePair = (labelRect, inputRect) => {
    ...        const labelCenterX = labelRect.left + (labelRect.width / 2);
    ...        const labelCenterY = labelRect.top + (labelRect.height / 2);
    ...        const inputCenterX = inputRect.left + (inputRect.width / 2);
    ...        const inputCenterY = inputRect.top + (inputRect.height / 2);
    ...        const dx = inputCenterX - labelCenterX;
    ...        const dy = inputCenterY - labelCenterY;
    ...        const horizontalGap = Math.max(0, inputRect.left - labelRect.right, labelRect.left - inputRect.right);
    ...        const verticalGap = Math.max(0, inputRect.top - labelRect.bottom, labelRect.top - inputRect.bottom);
    ...        const rowOverlap = Math.max(0, Math.min(labelRect.bottom, inputRect.bottom) - Math.max(labelRect.top, inputRect.top));
    ...        const columnOverlap = Math.max(0, Math.min(labelRect.right, inputRect.right) - Math.max(labelRect.left, inputRect.left));
    ...        const sameRow = rowOverlap > Math.min(labelRect.height, inputRect.height) * 0.25;
    ...        const sameColumn = columnOverlap > Math.min(labelRect.width, inputRect.width) * 0.2;
    ...        const toRight = inputRect.left >= labelRect.left - 20;
    ...        const below = inputRect.top >= labelRect.top - 20;
    ...        const rightPattern = sameRow && toRight;
    ...        const belowPattern = sameColumn && below;
    ...        const patternPenalty = rightPattern || belowPattern ? 0 : 3000;
    ...        const leftSidePenalty = inputRect.right < labelRect.left - 10 ? 2500 : 0;
    ...        const abovePenalty = inputRect.bottom < labelRect.top - 10 ? 2500 : 0;
    ...        const gapPenalty = horizontalGap * 1.4 + verticalGap * 1.8;
    ...        const centerPenalty = Math.abs(dx) * 0.35 + Math.abs(dy) * 0.35;
    ...        const rowBonus = rightPattern ? -120 : 0;
    ...        const columnBonus = belowPattern ? -90 : 0;
    ...        return patternPenalty + leftSidePenalty + abovePenalty + gapPenalty + centerPenalty + rowBonus + columnBonus;
    ...    };
    ...    const costMatrix = labels.map(({ rect }) => inputRects.map(({ rect: inputRect }) => scorePair(rect, inputRect)));
    ...    let bestAssignment = null;
    ...    let bestScore = Number.POSITIVE_INFINITY;
    ...    const search = (labelIndex, usedInputs, runningScore, assignment) => {
    ...        if (labelIndex === labels.length) {
    ...            if (runningScore < bestScore) {
    ...                bestScore = runningScore;
    ...                bestAssignment = assignment.slice();
    ...            }
    ...            return;
    ...        }
    ...        if (runningScore >= bestScore) return;
    ...        const options = inputRects
    ...            .map((_, inputIndex) => ({ inputIndex, score: costMatrix[labelIndex][inputIndex] }))
    ...            .filter(({ inputIndex }) => !usedInputs.has(inputIndex))
    ...            .sort((a, b) => a.score - b.score);
    ...        for (const { inputIndex, score } of options) {
    ...            usedInputs.add(inputIndex);
    ...            assignment[labelIndex] = inputIndex;
    ...            search(labelIndex + 1, usedInputs, runningScore + score, assignment);
    ...            usedInputs.delete(inputIndex);
    ...        }
    ...    };
    ...    search(0, new Set(), 0, []);
    ...    if (!bestAssignment) return null;
    ...    window.__rf_input_map = {};
    ...    window.__rf_input_map_signature = signature;
    ...    labels.forEach((item, index) => {
    ...        window.__rf_input_map[item.name] = inputRects[bestAssignment[index]].input;
    ...    });
    ...    return window.__rf_input_map[label] || null;
    ${element}=    Execute Javascript    ${script}    ARGUMENTS    ${label}
    IF    $element is None
        Fail    No visible input found for label '${label}'
    END
    RETURN    ${element}

Get Visible Button
    [Arguments]    ${label}
    ${script}=    Catenate    SEPARATOR=\n
    ...    const label = arguments[0];
    ...    const normalize = (value) => (value || '').replace(/\\s+/g, ' ').trim();
    ...    const isVisible = (el) => {
    ...        if (!el) return false;
    ...        const style = window.getComputedStyle(el);
    ...        const rect = el.getBoundingClientRect();
    ...        return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
    ...    };
    ...    const getPriority = (el) => {
    ...        const tag = (el.tagName || '').toLowerCase();
    ...        const type = (el.type || '').toLowerCase();
    ...        if (tag === 'button') return 0;
    ...        if (tag === 'input' && ['button', 'submit'].includes(type)) return 1;
    ...        if (tag === 'a' || (el.getAttribute('role') || '').toLowerCase() === 'button') return 2;
    ...        if ((el.className || '').toString().includes('Button')) return 3;
    ...        return 4;
    ...    };
    ...    const candidates = Array.from(document.querySelectorAll('button, a, input[type="button"], input[type="submit"], [role="button"], div, span'))
    ...        .filter((el) => {
    ...            const text = normalize(el.textContent || el.value);
    ...            return isVisible(el) && text === label;
    ...        })
    ...        .map((el) => {
    ...            const rect = el.getBoundingClientRect();
    ...            return { el, rect, priority: getPriority(el), area: rect.width * rect.height };
    ...        })
    ...        .sort((a, b) => (a.priority - b.priority) || (a.area - b.area) || (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
    ...    return candidates.length ? candidates[0].el : null;
    ${element}=    Execute Javascript    ${script}    ARGUMENTS    ${label}
    IF    $element is None
        Fail    No visible button found for label '${label}'
    END
    RETURN    ${element}

Get Button For Label
    [Arguments]    ${label}
    ${script}=    Catenate    SEPARATOR=\n
    ...    const label = arguments[0];
    ...    const normalize = (value) => (value || '').replace(/\\s+/g, ' ').trim();
    ...    const isPresent = (el) => {
    ...        if (!el) return false;
    ...        const style = window.getComputedStyle(el);
    ...        const rect = el.getBoundingClientRect();
    ...        return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
    ...    };
    ...    const getPriority = (el) => {
    ...        const tag = (el.tagName || '').toLowerCase();
    ...        const type = (el.type || '').toLowerCase();
    ...        if (tag === 'button') return 0;
    ...        if (tag === 'input' && ['button', 'submit'].includes(type)) return 1;
    ...        if (tag === 'a' || (el.getAttribute('role') || '').toLowerCase() === 'button') return 2;
    ...        if ((el.className || '').toString().includes('Button')) return 3;
    ...        return 4;
    ...    };
    ...    const candidates = Array.from(document.querySelectorAll('button, a, input[type="button"], input[type="submit"], [role="button"], div, span'))
    ...        .filter((el) => {
    ...            const text = normalize(el.textContent || el.value);
    ...            return isPresent(el) && text === label;
    ...        })
    ...        .map((el) => {
    ...            const rect = el.getBoundingClientRect();
    ...            return { el, rect, priority: getPriority(el), area: rect.width * rect.height };
    ...        })
    ...        .sort((a, b) => (a.priority - b.priority) || (a.area - b.area) || (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
    ...    return candidates.length ? candidates[0].el : null;
    ${element}=    Execute Javascript    ${script}    ARGUMENTS    ${label}
    IF    $element is None
        Fail    No button found for label '${label}'
    END
    RETURN    ${element}

Click Visible Element
    [Arguments]    ${element}
    Call Method    ${element}    click

Click Start Element
    [Arguments]    ${element}
    Run Keyword And Ignore Error    Call Method    ${element}    click
    Execute Javascript    arguments[0].click();    ARGUMENTS    ${element}
    Execute Javascript
    ...    const el = arguments[0];
    ...    const rect = el.getBoundingClientRect();
    ...    const clientX = rect.left + (rect.width / 2);
    ...    const clientY = rect.top + (rect.height / 2);
    ...    for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
    ...        el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX, clientY }));
    ...    }
    ...    ARGUMENTS
    ...    ${element}

Double Click Start Element
    [Arguments]    ${element}
    Click Start Element    ${element}
    Sleep    300ms
    Click Start Element    ${element}

Input By Label
    [Arguments]    ${label}    ${value}
    ${element}=    Wait Until Keyword Succeeds    6x    300ms    Get Visible Input For Label    ${label}
    Execute Javascript    arguments[0].scrollIntoView({block: 'center'});    ARGUMENTS    ${element}
    Clear Element Text    ${element}
    Input Text    ${element}    ${value}

Fill Form Dynamically
    [Arguments]    ${row}    ${is_last_row}=${False}
    Input By Label    Company Name    ${row["company_name"]}
    Input By Label    Address    ${row["company_address"]}
    Input By Label    EIN    ${row["employer_identification_number"]}
    Input By Label    Sector    ${row["sector"]}
    Input By Label    Automation Tool    ${row["automation_tool"]}
    Input By Label    Annual Saving    ${row["annual_automation_saving"]}
    Input By Label    Date    ${row["date_of_first_project"]}

    Click Submit Button
    IF    ${is_last_row}
        Wait For Completion Page
        RETURN
    END
    Wait Until Keyword Succeeds    10x    500ms    Get Visible Input For Label    Company Name

*** Test Cases ***
Automation Challenge Test
    ${data}=    Read Excel    ${EXCEL}
    ${row_count}=    Get Length    ${data}
    Activate Challenge

    FOR    ${index}    ${row}    IN ENUMERATE    @{data}
        ${is_last_row}=    Evaluate    ${index} == ${row_count} - 1
        Fill Form Dynamically    ${row}    ${is_last_row}
    END

    Log    Completed all entries
