const msalConfig = {
    appId       : import.meta.env.VITE_MS_CLIENT_ID,
    redirectUri : import.meta.env.VITE_MS_REDIRECT_URI,
    scopes      : [
        'user.read',
        'mailboxsettings.read',
        'calendars.readwrite'
    ]
};

export default msalConfig;