import * as React from 'react';
import ReactDOM from 'react-dom';
import $ from 'jquery';
import { Button, ButtonType, IButtonProps, Label } from 'office-ui-fabric-react/lib/index';

require('./goToSPFolder.html');
require('../node_modules/office-ui-fabric-react/dist/css/fabric.css');

var gotoSPFolder = window.gotoSPFolder || {};
gotoSPFolder.WebAPI = gotoSPFolder.WebAPI || {};
var goButton;

class GoButton extends React.Component<IButtonProps, {}> {
    constructor(props) {
        super(props);
        this.state = {
            description:'',
            disabled : false
        }
        this.updateDescription = this.updateDescription.bind(this);
    }
    updateDescription(desc){
        this.setState({description:desc});
    }
    updateDisabled(dis){
        this.setState({disabled: dis});
    }
    render() {
        return (
            <Button
                onClick={goToEntity}
                buttonType={ButtonType.compound}
                description={this.state.description}
                disabled = {this.state.disabled}
                icon='Folder'>
                Go To SharePoint Folders
                </Button>
        );
    }
}

$(document).ready(() => {

    goButton = ReactDOM.render(
    <GoButton />,
    document.getElementById('root')
    );

    getEntity(getURL()).then((data) => {
        if (data['@odata.etag']){
        goButton.updateDescription(data.new_name);
        goButton.updateDisabled(false);
        // console.log('GotData')
        // console.log(data);
        }else{
        goButton.updateDescription('');
        goButton.updateDisabled(true);
        console.log('GotNoData')
        }
    
    })
    .fail(() => {
        console.log('Failed ajax;')
    });
}
);

function goToEntity() {
    var parts = {};
    parts.query = queryStringtoObject(window.location.search);
    getEntity(getURL())
        .then((data) => {
            parts.entity = data;
            return getAccount(data);
        })
        .fail((err) => {
            console.log('AccountNotFound');
        })
        .then((data) => {
            parts.account = data;
            parts.folderURL = 'https://wagstaffinc.sharepoint.com/sites/CRM/SitePages/handelSPFolder.aspx?folder=/sites/CRM/account/' + sanatizePunctuation(parts.account.name) + '/' + parts.query.typename + '/' + sanatizePunctuation(parts.entity.new_name) + '&ProjectID=' + parts.entity.new_projectid;
            window.open(parts.folderURL);
        }).fail((err) => {
            console.log('foldernnotfound');
        })

}

function queryStringtoObject(search) {
    var p = {};
    var q = decodeURIComponent(search).replace('?', '');
    var vars = q.split('&');
    vars.forEach(function (element) {
        var name = element.split('=')[0];
        var value = decodeURI(element.split('=')[1]);
        if (value.indexOf('?') > 0) {
            queryStringtoObject(value);
        } else {
            p[name] = value;
        }
    }, this);

    return p;
}
function getURL() {
    var url = ''
    var q = queryStringtoObject(window.location.search);
    if (q.typename) {
        var win = window.location;
        switch (q.typename) {
            case 'new_projectactivities': {
                q.typename = 'new_projectactivitieses';
                break;
            }
            case 'account': {
                q.typename = 'accounts';
                break;
            }
            case 'opportunity': {
                q.typename = 'opportunities';
                break;
            }
            default: break;
        }
        url = win.protocol + "//" + win.hostname + "/api/data/v8.1/" + q.typename + "(" + q.id.replace("{", "").replace("}", "") + ")";
    } else {
        console.log('TypeNotFound');
    }
    return url;
}

function getEntity(_url = getURL()) {
    return $.ajax(
        {
            url: _url,
            method: 'GET',
            dataType: 'json',
        }
    );
}
function getAccount(data) {
    var win = window.location;
    var _url = win.protocol + "//" + win.hostname + "/api/data/v8.1/accounts(" + data._new_account_value + ")";
    return $.ajax({
        url: _url,
        method: 'GET',
        dataType: 'json'
    })
}

function sanatizePunctuation(str) {
    var t = str
        .replace(/\//g, '-')
        .replace(/\</g, '-')
        .replace(/\>/g, '-')
        .replace(/\./g, '-')
        .replace(/\?/g, '-')
        .replace(/\\/g, '-')
        .replace(/\,/g, '-')
        .replace(/\;/g, '-')
        .replace(/\:/g, '-')
        .replace(/\|/g, '-')
        .replace(/\!/g, '-')
        .replace(/\@/g, '-')
        .replace(/\#/g, '-')
        .replace(/\$/g, '-')
        .replace(/\%/g, '-')
        .replace(/\^/g, '-')
        .replace(/\&/g, '-')
        .replace(/\*/g, '-')
        .replace(/\_/g, '-')
        .replace(/\+/g, '-')
        .replace(/\=/g, '-')
    return t;


}
