var profilesArray = [];
var prefsObj = [];
var profileJSONuri = "json/profiles.json";
var prefsJSONuri = "json/prefs.json";


Office.initialize = function (reason) {
    $(document).ready(function () {
        if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
            console.log('Sorry. The add-in uses Word.js APIs version 1.3, that are not available in your version of Office.');
        };

        // INITIALIZE
        readFetchPreferences();
        
        $('#home-select-profile').focus();
        
        // UI
        $('.profile-buttons .ms-Button').click(function () {
            var selectedTab = $(this).attr('data-target');
            $('.profile-buttons').css('display', 'none');
            $('.profile-sections .profile-section').fadeOut(0, function () {
                $(selectedTab).css('display', 'block');
            });
        });

        $('.back-button').click(function () {
            $('.profile-sections .profile-section').fadeOut(0, function () {
                $('.profile-buttons').css('display', 'block');
            });
        });

        $('.ms-Pivot-link:nth-of-type(2)').click(function () {
            $('.profile-sections .profile-section').fadeOut(0, function () {
                $('.profile-buttons').css('display', 'block');
            });
        });

        // APPLY PROFILE LIST CHANGE
        $('#home-select-profile').change(function (e) {
            var profileNo = parseInt($('#home-select-profile option:selected').attr('value'));
            fillApplyList(profileNo);
        });

        // MODIFY PROFILE LIST CHANGE 
        $('#modify-select-profile').change(function (e) {
            // var profileNo = parseInt($('#modify-select-profile option:selected').attr('value'));
            fillModifyChecks();
        });

        // APPLY PROFILE BUTTON CLICK
        $('#apply-profile').click(function () {
            
        });

        // CREATE PROFILE BUTTON CLICK
        $('#btn-new-profile').click(function () {
            addProfile();
        });

        // MODIFY PROFILE BUTTON CLICK
        $('#btn-edit-profile').click(function () {
            modifyProfile();
        });

        // DELETE PROFILE BUTTON CLICK
        $('body').on('click', '.btn-deleteprofile', function () {
            var profileToDelete = parseInt($(this).parent().parent().attr('value'));
            removeProfile(profileToDelete);
        });

        // APPLY PREFERENCE BUTTON CLICK
        $('#apply-profile').click(function(){
            applyProfile();
        })
    });
};

// LOAD PREFERENCES FROM JSON, STORE THEM IN ARRAY AND FETCH THEM TO UI
function readFetchPreferences() {
    $.getJSON(prefsJSONuri, function (data) {
        prefsObj = data;
        prefsObj.forEach(function (pref) {
            var prefDisplayadd = '<div><label class="ms-Label prefLabel" title="' + pref.description + '">' + pref.title + ' :&nbsp;</label>';
            var prefDisplaymod = '<div><label class="ms-Label prefLabel" title="' + pref.description + '">' + pref.title + ' :&nbsp;</label>';
            for (let index = 0; index < pref.values.length; index++) {
                if (index == 0) {
                    prefDisplayadd = prefDisplayadd + '<label class="prefvalueLabel" title="' + pref.values[index].description + '"><input type="radio" name="add' + pref.name + '" value="' + index + '" checked />' + pref.values[index].title + '</label>'
                    prefDisplaymod = prefDisplaymod + '<label class="prefvalueLabel" title="' + pref.values[index].description + '"><input type="radio" name="mod' + pref.name + '" value="' + index + '" checked />' + pref.values[index].title + '</label>'
                } else {
                    prefDisplayadd = prefDisplayadd + '<label class="prefvalueLabel" title="' + pref.values[index].description + '"><input type="radio" name="add' + pref.name + '" value="' + index + '"/>' + pref.values[index].title + '</label>'
                    prefDisplaymod = prefDisplaymod + '<label class="prefvalueLabel" title="' + pref.values[index].description + '"><input type="radio" name="mod' + pref.name + '" value="' + index + '"/>' + pref.values[index].title + '</label>'
                }
            }
            prefDisplayadd = prefDisplayadd + '</div>'
            prefDisplaymod = prefDisplaymod + '</div>'
            $('.profile-preferences-section.add-section').append(prefDisplayadd);
            $('.profile-preferences-section.modify-section').append(prefDisplaymod);
            readFetchProfiles();
        });
    });
}

// TO READ PROFILES FROM JSON, STORE THEM IN ARRAY AND FETCH THEM TO UI 
function readFetchProfiles() {
    profilesArray = [];
    $.getJSON(profileJSONuri, function (data) {
        profilesArray = data;
        $('#home-select-profile').html("");
        $('#modify-select-profile').html("");
        $('.delete-profiles-table tbody').html("");

        for (var i = 0; i < profilesArray.length; i++) {
            $('#home-select-profile').append(
                '<option value=' + i + '>' + profilesArray[i].profilename + '</option>'
            );
            if (profilesArray[i].type == "user") {
                $('#modify-select-profile').append(
                    '<option value=' + i + '>' + profilesArray[i].profilename + '</option>'
                );
                $('.delete-profiles-table tbody').append(
                    '<tr value=' + i + '><td>' + profilesArray[i].profilename + '</td><td><button class="ms-Button btn-deleteprofile" title="Delete Profile">Delete</button></td></tr>'
                );
            }
        };
        fillApplyList(0);
        fillModifyChecks();
    });
};

// FILL THE APPLY PROFILE DETAILS WITH CORRESPONDING PROFILE PREFERENCES
function fillApplyList(profileNo) {
    var rawPrefs = profilesArray[profileNo].prefs;
    var prefs = rawPrefs.split("");
    var dataTarget = '.profile-preferences-section.apply-section';
    $(dataTarget).html("");
    for (var index = 0; index < prefsObj.length; index++) {
        var prefdisplay = '<label class="ms-Label"><span class="preference-title" title="' + prefsObj[index].description + '">' + prefsObj[index].title + ' : </span><span class="preference-value" title="' + prefsObj[index].values[prefs[index]].description + '">' + prefsObj[index].values[prefs[index]].title + '</span></label>';
        $(dataTarget).append(prefdisplay);
    };
};

// FILL THE MODIFY PROFILE DETAILS WITH CORRESPONDING PROFILE PREFERENCES
function fillModifyChecks() {
    var profileNo = parseInt($('#modify-select-profile option:selected').attr('value'));
    var rawPrefs = profilesArray[profileNo].prefs;
    var prefs = rawPrefs.split("");
    var dataTarget = '.profile-preferences-section.modify-section';
    for (var index = 0; index < prefsObj.length; index++) {
        var radio = 'input[name=mod' + prefsObj[index].name + ']'
        $(dataTarget).find(radio).eq(prefs[index]).prop("checked", true);
    };
};

// NOTIFICATION AND ERROR MESSAGES ALERT
function notifyMessage() {
    $('.notification-message').fadeIn(500, function () {
        $('.notification-message').fadeOut(3000)
    });
};

function errorMessage() {
    $('.error-message').fadeIn(500, function () {
        $('.error-message').fadeOut(3000)
    });
};

// TO ADD/EDIT A PROFILE
function addProfile() {
    var profName = $('#txt-newprofile-name').val();
    if (profName != "") {
        var profilePrefs = "";
        var newProfile = {
            "profilename": profName,
            "type": "user",
            "prefs": ""
        };

        for (var a = 0; a < prefsObj.length; a++) {
            var radio = 'input[name=add' + prefsObj[a].name + ']:checked'
            var prefValue = $('.profile-preferences-section.add-section').find(radio).attr('value');
            profilePrefs += prefValue;
        };

        newProfile.prefs = profilePrefs;
        profilesArray.push(newProfile);
        $.ajax({
            url: '/',
            type: "POST",
            data: {profiles : profilesArray},
            success: function (data, textStatus, jqXHR) {
                readFetchProfiles();
                $('#txt-newprofile-name').val('');
                var divs = $('.profile-preferences-section.add-section').find('div');
                for (let i=0; i<divs.length;i++){
                    divs.eq(i).find('input[type=radio]').eq(0).prop('checked',true);
                };
            }
        });
        notifyMessage();
    } else {
        errorMessage();
    };
};

// TO REMOVE A PROFILE
function removeProfile(profileNo) {
    profilesArray.splice(profileNo, 1);
        $.ajax({
            url: '/',
            type: "POST",
            data: {profiles : profilesArray},
            success: function (data, textStatus, jqXHR) {
                readFetchProfiles();
            }
        });
}

// TO MODIFY A PROFILE
function modifyProfile() {
    var profNo = parseInt($('#modify-select-profile option:checked').attr('value'));
    if (profNo >= 0) {
        var profilePrefs = "";

        for (var a = 0; a < prefsObj.length; a++) {
            var radio = 'input[name=mod' + prefsObj[a].name + ']:checked';
            var prefValue = $('.profile-preferences-section.modify-section').find(radio).attr('value');
            profilePrefs += prefValue;
        };

        profilesArray[profNo].prefs = profilePrefs;
        $.ajax({
            url: '/',
            type: "POST",
            data: {profiles : profilesArray},
            success: function (data, textStatus, jqXHR) {
                readFetchProfiles();
            }
        });
        notifyMessage();
    } else {
        errorMessage();
    };
};

// TO APPLY A PROFILE
function applyProfile(){
    var profNo = parseInt($('#home-select-profile option:checked').attr('value'));

    if (profNo >= 0){
        $("#apply-profile").attr("disabled", "disabled");
        $("#apply-profile span").text("Applying...");
        var selectedPrefs = profilesArray[profNo].prefs.slice("");
        applyFontPreference(selectedPrefs[1], function(){
            console.log('Font Preference Applied')
            applyAlignmentPreference(selectedPrefs[2], function(){
                console.log('Alignment Preference Applied')
                applySentenceSpacingPreference(selectedPrefs[0], function(){
                    console.log('Spacing Preference Applied')
                    // applycnparPreference(selectedPrefs[4], function(){
                        console.log('Paranthesis Preference Applied')
                        applyCnformatPreference(selectedPrefs[3],function(){
                            console.log('Case Name Preference Applied')
                                $("#apply-profile span").text("Apply Profile Preferences");
                                $("#apply-profile").removeAttr("disabled");
                                notifyMessage();
                        });
                    // });
                });    
            });
        });
    } else {
        errorMessage();
    };
};

function applySentenceSpacingPreference(optionNo, callback){
    Word.run(function (context) {
        // List Of Exceptions in Search Results
        var exceptions = [".1",".2",".3",".4",".5",".6",".7",".8",".9",".0","2d","3d",". id",".  id",". corp",". co", ".com",".uk",".us",".gov",".org",".ca",".edu",".html",".asp",".php",".io",".ed"]
        
        var rule = prefsObj[0].values[optionNo].rule;
        var paras = context.document.body.paragraphs;
        context.load(paras, 'text')
        return context.sync().then(function () {
            var res = [];
            paras.items.forEach(function(para){
                if (para.text.length >100) {
                    var searchTermsComplete = [];
                    var searchTerms = [];
                    var revStr = para.text.split('').reverse().join('');
                    var reverseRexexp = /\b\w+\b(“?|\(?)(\s*)(\.|”\.)(?!rj )(?!rM )(?!sm )(?!srm )(?!rd )(?!rs )(?!ssim )(?!de |de\.)(?!di |di\.)(?!tc |tc\.)(?!rtpr |rtpr\.)(?!v )(?!lac)(?!ppa |ppa\.)(?!qse )(?!on )(?!xe )(?!e\.i)(?!m\.a)(?!m\.p)(?!\w+@)(?!cni )(?!oc )(?!proc)(?!tsid)(?!www)(?!Y\.N)(?!C\.S)(?!S\.U)(?!K\.U)(?![A-Z] )(?![A-Z]\.)(?![A-Z]$)(?![A-Z]\()(?!naj\()(?!bef\()(?!ram\()(?!rpa\()(?!yam\()(?!nuj\()(?!luj\()(?!gua\()(?!tpes\()(?!pes\()(?!tco\()(?!von\()(?!ced\()(?!tra )/gi
                    var myArray;
                    while ((myArray = reverseRexexp.exec(revStr)) !== null) {
                        var result = myArray[0];
                        searchTermsComplete.push(result.split('').reverse().join('').toLocaleLowerCase())
                    }

                    let cur;
                    searchTermsComplete.sort().forEach(function(x) {
                        if (!x.startsWith(cur)) {
                            searchTerms.push(x);
                            cur = x
                        }
                    })

                    // Push search Results into res array
                    for (var index = 0; index < searchTerms.length; index++) {
                        res.push(para.search(searchTerms[index], {matchWildcards: false}));
                    };
                }
            })
            for (let index = 0; index < res.length; index++) {
                res[index].load('font/italic, font/underline, text, hyperlink');   
            }
            return context.sync().then(function(){
                for (let index = 0; index < res.length; index++) {
                    
                    res[index].items.forEach(function(r){
                        if (!exceptions.some(function(exep) {return r.text.toLowerCase().includes(exep)})){
                            if (!r.hyperlink && !r.text.startsWith(". at")){
                                // console.log(r.text);
                                if (r.text.startsWith('.”')){
                                    var newText = ".”" + rule + r.text.slice(2).trim();
                                    r.insertText(newText, "Replace");
                                } else {
                                    var newText = "." + rule + r.text.slice(1).trim();
                                    r.insertText(newText, "Replace");
                                }
                                r.font.set({
                                    italic: false,
                                    underline: "None"
                                });
                            }
                        }                        
                    })
                }
                return context.sync().then(function(){
                    callback();
                });
            })
        });
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });
}

function applyFontPreference(optionNo, callback){
    Word.run(function (context) {
        let rule = prefsObj[1].values[optionNo].rule;
        let paras = context.document.body.paragraphs;
        let header = context.document.sections.getFirst().getHeader('Primary'); 
        let footer = context.document.sections.getFirst().getFooter('Primary');
        paras.load('font/name')
        header.load('font/name')
        footer.load('font/name')
        return context.sync().then(function(){
            // if  (body.font.name != rule){
                paras.items.forEach(function(para){
                    para.font.name = rule;
                })
            // }
            // if (footer.font.name != rule){
                footer.font.name = rule;
            // }
            // if(header.font.name != rule){
                header.font.name = rule; 
            // }
            return context.sync().then(callback())
        })
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });
}

function applyAlignmentPreference(optionNo, callback){
    Word.run(function (context) {
        var rule = prefsObj[2].values[optionNo].rule;
        var paras = context.document.body.paragraphs;
        paras.load("alignment");
        return context.sync().then(function () {            
            paras.items.forEach(function (para) {
                if (para.alignment != "Centered" && para.alignment != rule){
                    para.alignment = rule;
                }
            });
            return context.sync().then(callback())
        });
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });;
}

function applyCnformatPreference(optionNo, callback){
    Word.run(function (context) {
        let body = context.document.body;
        body.load('text');
        return context.sync().then(function () {
            var searchTerms = findAllCases(body.text);
            var res = [];
            for (let i = 0; i < searchTerms.length; i++) {
                res[i] = context.document.body.search(searchTerms[i], {matchWildcards: false});
            };
            res.push(body.search("<[iI]d>.",{matchWildcards: true}));
            res.push(body.search("<[iI]bid>.",{matchWildcards: true}));
            for (let i =0; i<res.length; i++){
                res[i].load('font/italic, font/underline, text');
            };
            return context.sync().then(function () {
                var results = [];
                for (let b=0; b<res.length;b++){
                    res[b].items.forEach(function(r){
                        results.push(r);
                    })
                }
                var newResults= [];
                // var newResultsShort = [];
                // console.log(results.length)
                for (let i = 0; i < results.length; i++) {
                    // if(results[i].text.toLowerCase().includes('supra')){
                        // newResultsShort.push(results[i].search("<*>", { matchWildcards: true }));
                    // } else {
                        newResults.push(results[i]);
                    // }
                };

                for (let i=0; i < newResults.length; i++){
                    if (prefsObj[3].values[optionNo].rule == "italic"){
                        newResults[i].font.set({
                            italic: true,
                            underline: "None"
                        });
                    } else if (prefsObj[3].values[optionNo].rule == "underline"){
                        newResults[i].font.set({
                            underline: "Single",
                            italic: false
                        });
                    };
                };
                // for (let i = 0; i < newResultsShort.length; i++) {
                //     newResultsShort[i].load('font/italic, font/underline');
                // }
                // return context.sync().then(function(){
                //     for (let i = 0; i < newResultsShort.length; i++) {
                //         newResultsShort[i].items.forEach(function(r){
                //             if (prefsObj[3].values[optionNo].rule == "italic"){
                //                 r.font.set({
                //                     italic: true,
                //                     underline: "None"
                //                 });
                //             } else if (prefsObj[3].values[optionNo].rule == "underline"){
                //                 r.font.set({
                //                     underline: "Single",
                //                     italic: false
                //                 });
                //             };      
                //         })                  
                //     }
                    return context.sync().then(function(){
                        callback();
                    });
                // })
            });
        });
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });
};

function applycnparPreference(optionNo, callback){
    Word.run(function (context) {
        let paras = context.document.body.paragraphs;
        paras.load('text');
        return context.sync().then(function () {
            var rangeCollects = [];
            var rangeCollectsAround = [];
            paras.items.forEach(function(para){
                if (para.text.length > 100){
                    var searchTerms = findAllCasesComplete(para.text);
                    var searchTermsAround = findAllCasesCompleteAround(para.text);
                    // console.log("Search Terms Length")
                    // console.log(searchTerms.length)
                    // console.log(searchTermsAround.length)
                    // console.log(para.text)
                    if (searchTerms.length > 0){
                        for (let index = 0; index < searchTerms.length; index++) {
                            rangeCollects.push(para.search(searchTerms[index], {matchWildcards: false}));
                        }
                    }
                    if (searchTermsAround.length > 0){
                        for (let index = 0; index < searchTermsAround.length; index++) {
                            rangeCollectsAround.push(para.search(searchTermsAround[index], {matchWildcards: false}))   
                        }
                    }
                }
            })
            
            for (let index = 0; index < rangeCollects.length; index++) {
                rangeCollects[index].load('text, font/underline, font/italic');
            }
            for (let index = 0; index < rangeCollectsAround.length; index++) {
                rangeCollectsAround[index].load('text, font/underline, font/italic');
            }
            return context.sync().then(function(){
                // console.log("Range Collections Length")
                // console.log(rangeCollects.length)
                // console.log(rangeCollectsAround.length)
                var allRanges = [];
                var allRangesAround = [];
                for (let index = 0; index < rangeCollects.length; index++) {
                    rangeCollects[index].items.forEach(function(rc){
                        allRanges.push(rc);
                    })
                }
                for (let index = 0; index < rangeCollectsAround.length; index++) {
                    rangeCollectsAround[index].items.forEach(function(rc){
                        allRangesAround.push(rc);
                    })
                }
                var rule = prefsObj[4].values[optionNo].rule;
                // console.log("Ranges Length")
                // console.log(allRanges.length)
                // console.log(allRangesAround.length)
                // console.log("Ranges")
                // for (let i = 0; i < allRanges.length; i++) {
                //     console.log(allRanges[i].text)
                // }
                // console.log("Ranges Around")
                // for (let i = 0; i < allRangesAround.length; i++) {
                //     console.log(allRangesAround[i].text)
                // }
                // if (allRanges.length == allRangesAround.length){
                    for (let index = 0; index < allRanges.length; index++) {
                    if (rule == "yes"){
                        if (allRangesAround[index]){
                            if(allRangesAround[index].text[0] != "("){
                                allRanges[index].insertText("(", Word.InsertLocation.start)
                            }
                            if(allRangesAround[index].text[allRangesAround[index].text.length-1] != ")"){
                                allRanges[index].insertText(")", Word.InsertLocation.end)
                            }
                        } else {
                            allRanges[index].insertText("(", Word.InsertLocation.start)
                            allRanges[index].insertText(")", Word.InsertLocation.end)
                        }
                        allRanges[index].font.set({
                            italic: false,
                            underline: "None"
                        })
                    } else if (rule == "no"){
                        if (allRangesAround[index].text[0] == "(" && allRangesAround[index].text[allRangesAround[index].text.length-1] == ")"){
                            allRangesAround[index].insertText(allRanges[index].text, Word.InsertLocation.replace);
                            allRangesAround[index].font.set({
                                italic: false,
                                underline: "None"
                            })
                        }
                    }
                // }
            }
                return context.sync().then(function(){
                    callback();
                });
            });            
        });
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });
};

function findAllCasesComplete(str){
    // var regex1 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor |e\.g\., |see |See |e\.g\.|also )*v\.\s*([A-Z][a-zA-Z,\.’]+\s*|of |(\(\d*\))|Cal\.|App\.|,|-|–|\b\d*\b|\b\d\w{1,2}\b| {1})*/g;
    var regex1 = /(([A-Z][A-Za-z,\.’]+ |in |of |a |minor |e\.g\., |see |See |e\.g\.|also )*v\.\s*([A-Z][a-zA-Z,\.’]+\s*|of |(\(\d*\))|Cal\.|App\.|,|-|–|\b\d*\b|\b\d\w{1,2}\b| {1}|;)*){1,}/g;
    // var regex2 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor |also |see |e\.g\.,| |e\.g\.)*supra, (\b\d+\b|Cal\.|at|App\.|\d\w{1,}|–| )*/g;
    var regex2 = /(([A-Z][A-Za-z,\.’]+ |in |of |a |minor |also |see |e\.g\.,| |e\.g\.)*supra, (\b\d+\b|Cal\.|at|App\.|\d\w{1,}|–| |;)*){1,}/g;
    // var regex3 = /\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|“[A-Z]”|[A-Z])*/g;
    // var regex3 = /\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|“[A-Z]”|[A-Z] )*/g;
    var regex3 = /(\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|;|“[A-Z]”|[A-Z] )*){1,}/g;
    var regex4 = /\b[iI]bid\b\./g;
    var Terms1 = findCaseNamesComplete(regex1, str);
    var Terms2 = findCaseNamesComplete(regex2, str);
    var Terms3 = findCaseNamesComplete(regex3, str);
    var Terms4 = findCaseNamesComplete(regex4, str);
    var Terms = [];
    Terms1.concat(Terms2, Terms3, Terms4).forEach(function(item){
        if (Terms.indexOf(item) == -1){
            Terms.push(item);
        }
    });
    return Terms;
}

function findAllCasesCompleteAround(str){
    // var regex1 = /.?([A-Z][A-Za-z,\.’]+ |in |of |a |minor |e\.g\., |see |See |e\.g\.|also )*v\.\s*([A-Z][a-zA-Z,\.’]+\s*|of |(\(\d*\))|Cal\.|App\.|,|-|–|\b\d*\b|\b\d\w{1,2}\b| {1})*.?/g;
    var regex1 = /.?(([A-Z][A-Za-z,\.’]+ |in |of |a |minor |e\.g\., |see |See |e\.g\.|also )*v\.\s*([A-Z][a-zA-Z,\.’]+\s*|of |(\(\d*\))|Cal\.|App\.|,|-|–|\b\d*\b|\b\d\w{1,2}\b| {1}|;)*){1,}.?/g;
    // var regex2 = /.?([A-Z][A-Za-z,\.’]+ |in |of |a |minor |also |see |e\.g\.,| |e\.g\.)*supra, (\b\d+\b|Cal\.|at|App\.|\d\w{1,}|–| )*.?/g;
    var regex2 = /.?(([A-Z][A-Za-z,\.’]+ |in |of |a |minor |also |see |e\.g\.,| |e\.g\.)*supra, (\b\d+\b|Cal\.|at|App\.|\d\w{1,}|–| |;)*){1,}.?/g;
    // var regex3 = /.?\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|“[A-Z]”|[A-Z])*.?/g;
    // var regex3 = /.?\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|“[A-Z]”|[A-Z] )*.?/g;
    var regex3 = /.?(\b[iI]d\b\.(at| |¶|\d|-|–|,|Ex\.|Exs\.|ex\.|exs\.|;|“[A-Z]”|[A-Z] )*){1,}.?/g;
    var regex4 = /.?\b[iI]bid\b\..?/g;
    var Terms1 = findCaseNamesComplete(regex1, str);
    var Terms2 = findCaseNamesComplete(regex2, str);
    var Terms3 = findCaseNamesComplete(regex3, str);
    var Terms4 = findCaseNamesComplete(regex4, str);
    var Terms = []

    Terms1.concat(Terms2, Terms3, Terms4).forEach(function(item){
        if (Terms.indexOf(item) == -1){
            Terms.push(item);
        }
    });
    return Terms;
}

function findCaseNamesComplete(regex, str){
    var searchTermsComplete = [];
    var myArray;
    while ((myArray = regex.exec(str)) !== null) {
        var result = myArray[0];
        searchTermsComplete.push(result)
    }
    for (let index = 0; index < searchTermsComplete.length; index++) {
        searchTermsComplete[index] = searchTermsComplete[index].trim();
        if(searchTermsComplete[index][searchTermsComplete[index].length-1] == ";"){
            searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1)
        }
        if(searchTermsComplete[index][searchTermsComplete[index].length-1] == ","){
            searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1)
        }
        if(searchTermsComplete[index].toLocaleLowerCase().startsWith("in ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(3, searchTermsComplete[index].length)
        }
        if(searchTermsComplete[index].startsWith("AAC. ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(5, searchTermsComplete[index].length)
        }
        searchTermsComplete[index] = searchTermsComplete[index].trim();
    }
    return searchTermsComplete;
}

function findAllCases(str){
    var regex1 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor )*v\.\s([A-Z][a-zA-Z,\.’]+(\s|\))|of )*/g;
    var regex2 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor )*supra,/g;
    var Terms1 = findCaseNames(regex1, str);
    var Terms2 = findCaseNames(regex2, str);
    var Terms = [];
    Terms1.concat(Terms2).forEach(function(item){
        if (Terms.indexOf(item) == -1){
            Terms.push(item);
        }
    });
    return Terms;
}

function findCaseNames(regex, str){
    var searchTermsComplete = [];
    var myArray;
    while ((myArray = regex.exec(str)) !== null) {
        var result = myArray[0];
        searchTermsComplete.push(result)
    }
    for (let index = 0; index < searchTermsComplete.length; index++) {
        searchTermsComplete[index] = searchTermsComplete[index].trim();
        // if(searchTermsComplete[index][searchTermsComplete[index].length-1] == ","){
            // searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1)
        // }
        if(searchTermsComplete[index].toLocaleLowerCase().includes("see ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(searchTermsComplete[index].toLocaleLowerCase().indexOf("see ") + 4, searchTermsComplete[index].length)
        }
        if(searchTermsComplete[index].toLocaleLowerCase().startsWith("in ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(3, searchTermsComplete[index].length)
        }
        if(searchTermsComplete[index][searchTermsComplete[index].length-1] ==")"){
            searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1);
        }
        // if(searchTermsComplete[index][searchTermsComplete[index].length-1] ==")"){
        //     searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1);
        // }
    }
    return searchTermsComplete;
}