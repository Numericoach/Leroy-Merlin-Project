module.exports = {
	"env": {
		"browser": true,
		"es2021": true,
		'googleappsscript/googleappsscript': true
	},
	"extends": [
		"eslint:recommended",
		"plugin:@typescript-eslint/recommended",
        "plugin:jsdoc/recommended"
	],
	"overrides": [
	],
	"parser": "@typescript-eslint/parser",
	"parserOptions": {
		"ecmaVersion": "latest",
		"sourceType": "module"
	},
	"plugins": [
		"@typescript-eslint",
		'googleappsscript',
	],
	'rules': {
		"camelcase": 1,
		"default-case": 1,
		"id-length": ["error", {"min": 2}],
		"max-depth": ["error", 3],
		"max-params": ["error", 4],
		"new-cap": 1,
		"no-magic-numbers": ["warn", {"ignoreDefaultValues": true, "ignoreClassFieldInitialValues": true}],
		"no-var": 2,
		"prefer-const": 1,
		"no-useless-constructor": 2,
		"jsdoc/require-param-type": 0,
		"jsdoc/require-returns-type": 0,
		"no-useless-return": 2,
		"yoda": 2,
		"prefer-arrow-callback": 2,
		"max-classes-per-file": ["error", 1],
		"no-cond-assign": 2,
		"no-unused-vars": 0,
		"@typescript-eslint/no-unused-vars": 0,
		"no-undef": 0,
		'jsdoc/no-undefined-types': 'off',
		"@typescript-eslint/no-namespace": 0,
	  },
}
