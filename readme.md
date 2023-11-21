# LSTM Forcasting SO2
pass

# Dependencies
- tensorflow with GPU
> https://www.tensorflow.org/install/pip?hl=th#windows-native
```bash
conda install -c conda-forge cudatoolkit=11.2 cudnn=8.1.0
# Anything above 2.10 is not supported on the GPU on Windows Native
python -m pip install "tensorflow<2.11"
# Verify the installation:
python -c "import tensorflow as tf; print(tf.config.list_physical_devices('GPU'))"
```

- Other dependencies
```bash
pip install -r requirements.txt
```