# Troubleshooting

## NumPy MKL library load failed on Jupyter Notebook (ipykernel)

1. Run systempropertiesadvanced.
1. Add CONDA_DLL_SEARCH_MODIFICATION_ENABLE=1.

### Reference

1. https://docs.conda.io/projects/conda/en/latest/user-guide/troubleshooting.html#numpy-mkl-library-load-failed
1. https://stackoverflow.com/questions/65888280/can-t-import-numpy-from-an-installed-jupyter-kernel
