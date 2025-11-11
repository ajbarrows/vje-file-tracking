# VJE Library File Organization and Monitoring

## Getting started

1.  Build Conda environment

```         
conda env create --name vje-ibrary --file=environment.yml
```

2.  [Set up OAuth2.0 token through Google's API](https://docs.iterative.ai/PyDrive2/quickstart/)

3.  Download `client_secrets.json` (named this way) and place in `./conf`

4.  Run make file to generate report

```         
make run
```
