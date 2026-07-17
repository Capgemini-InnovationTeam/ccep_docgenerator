import fioriTools from '@sap-ux/eslint-plugin-fiori-tools';

export default [
    {
        ignores: [
            "dist/**",
            "webapp/thirdparty/**"
        ]
    },
    ...fioriTools.configs.recommended
];