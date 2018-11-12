var profilesArray = [];
var prefsObj = [];
var profileJSONuri = "https://api.myjson.com/bins/18m48y";
var prefsJSONuri = "https://api.myjson.com/bins/vtrpu";

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
    console.log("Reading Preferences")
    $.getJSON(prefsJSONuri, function (data) {
        prefsObj = data;
        console.log("Preferences Read...")
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
        var profilesString = JSON.stringify(profilesArray);
        $.ajax({
            url: profileJSONuri,
            type: "PUT",
            data: profilesString,
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (data, textStatus, jqXHR) {
                readFetchProfiles();
                var profName = $('#txt-newprofile-name').val('');
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
    var profilesString = JSON.stringify(profilesArray);
    $.ajax({
        url: profileJSONuri,
        type: "PUT",
        data: profilesString,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
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
        var profilesString = JSON.stringify(profilesArray);
        $.ajax({
            url: profileJSONuri,
            type: "PUT",
            data: profilesString,
            contentType: "application/json; charset=utf-8",
            dataType: "json",
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
        var selectedPrefs = profilesArray[profNo].prefs.slice("");
        applySentenceSpacingPreference(selectedPrefs[0]);
        applyFontPreference(selectedPrefs[1]);
        applyAlignmentPreference(selectedPrefs[2]);
        applycnparPreference(selectedPrefs[4], function(){
            applyCnformatPreference(selectedPrefs[3]);
        });
        notifyMessage();
    } else {
        errorMessage();
    };
};

function applySentenceSpacingPreference(optionNo){
    Word.run(function (context) {
        var exceptions = [".1",".2",".3",".4",".5",".6",".7",".8",".9",".0","2d","3d", " ibid", ". at", ". id",".  id",". corp",". co", ".com",".uk",".us",".gov",".org",".ca",".edu",".html",".asp",".io",".ed"]
        var rule = prefsObj[0].values[optionNo].rule;
        
        var paras = context.document.body.paragraphs;
        paras.load('items');
        return context.sync().then(function () {
            var res = [];
            paras.items.forEach(function(para){
                var searchTermsComplete = [];
                var searchTerms = [];
                var revStr = para.text.split('').reverse().join('');
                var reverseRexexp = /\b\w+\b *\.(?!rj )(?!rM )(?!sm )(?!srm )(?!rd )(?!rs )(?!ssim )(?!de |de\.)(?!di |di\.)(?!tc |tc\.)(?!rtpr |rtpr\.)(?!v )(?!lac)(?!dibi )(?!ppa |ppa\.)(?!qse )(?!on )(?!xe )(?!e\.i)(?!m\.a)(?!m\.p)(?!\w+@)(?!cni )(?!oc )(?!www)(?!Y\.N)(?!C\.S)(?!S\.U)(?![A-Z] )(?![A-Z]\.)(?![A-Z]$)(?![A-Z]\()(?!naj\()(?!bef\()(?!ram\()(?!rpa\()(?!yam\()(?!nuj\()(?!luj\()(?!gua\()(?!tpes\()(?!tco\()(?!von\()(?!ced\()(?!tra )/gi
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

                for (var index = 0; index < searchTerms.length; index++) {
                    var rangeCollection = para.search(searchTerms[index]);
                    res.push(rangeCollection);
                };
            })
            res.forEach(function(r){
                r.load('items')
            })
            return context.sync().then(function(){
                for (let index = 0; index < res.length; index++) {
                    res[index].items.forEach(function(r){
                        if (!exceptions.some(function(exep) {return r.text.toLowerCase().includes(exep)})){
                            if (!r.hyperlink){
                                var newText = "." + rule + r.text.slice(1).trim();
                                r.insertText(newText, Word.InsertLocation.replace)
                            }
                        }
                    })
                }    
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

function applyFontPreference(optionNo){
    Word.run(function (context) {
        const body = context.document.body;
            body.font.set({
                name: prefsObj[1].values[optionNo].rule
            });
        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });
}

function applyAlignmentPreference(optionNo){
    Word.run(function (context) {
        var paras = context.document.body.paragraphs;
        paras.load("items/alignment");
        return context.sync().then(function () {
            paras.items.forEach(function (para) {
                if (para.alignment != "Centered"){
                    para.alignment = prefsObj[2].values[optionNo].rule;
                }
            });
        });
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        };
    });;
}

function applyCnformatPreference(optionNo){
    Word.run(function (context) {
        let body = context.document.body;
        body.load('text');
        return context.sync().then(function () {
            var searchTerms = findAllCases(body.text);
            var res = [];
            for (let i = 0; i < searchTerms.length; i++) {
                res[i] = context.document.body.search(searchTerms[i]);
            };
            res.push(context.document.body.search("<[iI]d.",{matchWildcards: true}));
            for (let i =0; i<res.length; i++){
                res[i].load('font');
            };
            return context.sync().then(function () {
                var results = [];
                for (let b=0; b<res.length;b++){
                    res[b].items.forEach(function(r){
                        results.push(r);
                    })
                }
                for (let i = 0; i < results.length; i++) {
                    if (prefsObj[3].values[optionNo].rule == "italic"){
                            results[i].font.set({
                                italic: true,
                                underline: "None"
                            });
                    } else if (prefsObj[3].values[optionNo].rule == "underline"){
                        results[i].font.set({
                            underline: "Single",
                            italic: false
                        });
                    };
                };
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
        let body = context.document.body;
        body.load('text');
        return context.sync().then(function () {
            var searchTerms = findAllCases(body.text);
            var searchTermsAround = [];
            searchTerms.forEach(function(s){
                searchTermsAround.push("?" + s + "?");
            })
            var res = [];
            var resAround = [];
            for (let i = 0; i < searchTerms.length; i++) {
                res[i] = context.document.body.search(searchTerms[i]);
                resAround[i] = context.document.body.search(searchTermsAround[i],{matchWildcards:true})
            }
            res.push(context.document.body.search("<[iI]d.",{matchWildcards: true}));
            resAround.push(context.document.body.search("?<[iI]d.?",{matchWildcards: true}))
            
            for (let i =0; i <res.length; i++){
                res[i].load('items');
                resAround[i].load('items');
            };
            return context.sync().then(function () {
                var results = [];
                var resultsAround = [];
                for (let b=0; b<res.length;b++){
                    res[b].items.forEach(function(r){
                        results.push(r);
                    });
                    resAround[b].items.forEach(function(r){
                        resultsAround.push(r);
                    });
                };
                for (let i = 0; i < results.length; i++) {
                    if (prefsObj[4].values[optionNo].rule == "yes"){
                        if (resultsAround[i].text[0] != "(" || resultsAround[i].text[resultsAround[i].text.length-1] != ")"){
                            results[i].insertText("(", Word.InsertLocation.start)
                            results[i].insertText(")", Word.InsertLocation.end)
                        };
                    } else if (prefsObj[4].values[optionNo].rule == "no"){
                        if (resultsAround[i].text[0] == "(" && resultsAround[i].text[resultsAround[i].text.length-1] == ")"){
                            resultsAround[i].insertText(results[i].text, Word.InsertLocation.replace);
                        };
                    };
                };
                callback();
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

function findAllCases(str){
    var regex1 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor )*v\.\s([A-Z][a-zA-Z,\.’]+(\s|\))|of )*/g;
    var regex2 = /([A-Z][A-Za-z,\.’]+ |in |of |a |minor )*supra/g;
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
        if(searchTermsComplete[index][searchTermsComplete[index].length-1] == ","){
            searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1)
        }
        if(searchTermsComplete[index].toLocaleLowerCase().includes("see ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(searchTermsComplete[index].toLocaleLowerCase().indexOf("see ") + 4, searchTermsComplete[index].length)
        }
        if(searchTermsComplete[index].toLocaleLowerCase().startsWith("in ")){
            searchTermsComplete[index] = searchTermsComplete[index].slice(3, searchTermsComplete[index].length)
        }
        if(searchTermsComplete[index][searchTermsComplete[index].length-1] ==")"){
            searchTermsComplete[index] = searchTermsComplete[index].slice(0, searchTermsComplete[index].length-1);
        }
    }
    return searchTermsComplete;
}