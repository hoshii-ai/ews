module github.com/hoshii-ai/ews

go 1.23.0

require (
	github.com/Azure/go-ntlmssp v0.0.0-20221128193559-754e69321358
	github.com/google/uuid v1.6.0
	github.com/pkg/errors v0.9.1
	github.com/stretchr/testify v1.4.0
)

require (
	github.com/davecgh/go-spew v1.1.0 // indirect
	github.com/pmezard/go-difflib v1.0.0 // indirect
	golang.org/x/crypto v0.0.0-20191206172530-e9b2fee46413 // indirect
	gopkg.in/yaml.v2 v2.2.2 // indirect
)

replace github.com/hoshii-ai/ews => ./

replace github.com/hoshii-ai/ews/utils => ./utils

replace github.com/hoshii-ai/ews/ewsutil => ./ewsutil
