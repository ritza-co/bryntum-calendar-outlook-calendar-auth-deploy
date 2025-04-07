import js from '@eslint/js';
import globals from 'globals';
import react from 'eslint-plugin-react';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';
import * as tseslint from 'typescript-eslint';

export default [
    { ignores : ['dist'] },
    ...tseslint.configs.recommended,
    {
        files           : ['**/*.{js,jsx,ts,tsx}'],
        languageOptions : {
            ecmaVersion   : 2020,
            globals       : globals.browser,
            parser        : tseslint.parser,
            parserOptions : {
                ecmaVersion  : 'latest',
                ecmaFeatures : { jsx : true },
                sourceType   : 'module'
            }
        },
        settings : { react : { version : '18.3' } },
        plugins  : {
            react,
            'react-hooks'   : reactHooks,
            'react-refresh' : reactRefresh
        },
        rules : {
            ...js.configs.recommended.rules,
            ...react.configs.recommended.rules,
            ...react.configs['jsx-runtime'].rules,
            ...reactHooks.configs.recommended.rules,
            'react/jsx-no-target-blank'            : 'off',
            'react-refresh/only-export-components' : [
                'warn',
                { allowConstantExport : true }
            ],
            'react/jsx-uses-react'     : 'off',
            'react/react-in-jsx-scope' : 'off',
            'accessor-pairs'           : 'off',
            'array-bracket-spacing'    : ['warn', 'never'],
            'array-callback-return'    : 'off',
            'brace-style'              : [
                'warn',
                'stroustrup',
                {
                    allowSingleLine : false
                }
            ],
            'comma-dangle'              : ['error', 'never'],
            'computed-property-spacing' : ['warn', 'never'],
            'dot-notation'              : 'off',
            eqeqeq                      : 'off',
            indent                      : [
                'error',
                4,
                {
                    SwitchCase   : 1,
                    ignoredNodes : ['TemplateLiteral']
                }
            ],
            'key-spacing' : [
                'warn',
                {
                    multiLine : {
                        afterColon  : true,
                        align       : 'colon',
                        beforeColon : true
                    },
                    singleLine : {
                        afterColon  : true,
                        beforeColon : true
                    }
                }
            ],
            'linebreak-style'   : ['error', 'unix'],
            'multiline-ternary' : 'off',
            'new-cap'           : [
                'warn',
                {
                    capIsNew   : false,
                    newIsCap   : false,
                    properties : false
                }
            ],
            'no-duplicate-imports'         : 'error',
            'no-extra-boolean-cast'        : 'off',
            'no-inner-declarations'        : 'off',
            'no-mixed-operators'           : 'off',
            'no-multi-spaces'              : 'off',
            'no-new-func'                  : 'off',
            'no-new-wrappers'              : 'off',
            'no-prototype-builtins'        : 'off',
            'no-restricted-globals'        : ['error', 'event', 'describe'],
            'no-return-assign'             : 'off',
            'no-trailing-spaces'           : 'warn',
            'no-unmodified-loop-condition' : 'off',
            'no-unused-expressions'        : 'off',
            'no-use-before-define'         : 'off',
            'node/no-callback-literal'     : 'off',
            'object-curly-spacing'         : ['warn', 'always'],
            'one-var'                      : 'off',
            'padded-blocks'                : 'off',
            'prefer-const'                 : 'warn',
            'prefer-promise-reject-errors' : 'off',
            'prefer-regex-literals'        : 'off',
            quotes                         : [
                'warn',
                'single',
                {
                    allowTemplateLiterals : true,
                    avoidEscape           : true
                }
            ],
            semi                           : ['error', 'always'],
            'space-before-function-paren'  : ['warn', 'never'],
            'spaced-comment'               : 'off',
            'standard/no-callback-literal' : 'off',
            'template-curly-spacing'       : 'off',
            yoda                           : [
                'error',
                'never',
                {
                    onlyEquality : true
                }
            ],
            'n/no-callback-literal' : 'off',
            'react/prop-types'      : 0
        }
    }
];
