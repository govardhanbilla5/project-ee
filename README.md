Code:.
│   aluminium_data.xlsx
│   best_price_model.pkl
│   input.xlsx
│   main.py
│   test.py
│   
├───app
│   │   __init__.py
│   │   
│   ├───api
│   │   │   __init__.py
│   │   │   
│   │   ├───v1
│   │   │   │   router.py
│   │   │   │   __init__.py
│   │   │   │   
│   │   │   ├───endpoints
│   │   │   │   │   forecast.py
│   │   │   │   │   forecast_batch.py
│   │   │   │   │   forecast_excel.py
│   │   │   │   │   forecast_ml.py
│   │   │   │   │   __init__.py
│   │           
│   ├───core
│   │   │   config.py
│   │   │   logging_config.py
│   │   │   __init__.py
│   │           
│   ├───data
│   │   │   base.py
│   │   │   excel_store.py
│   │   │   hardcoded_store.py
│   │   │   __init__.py
│   │           
│   ├───models
│   │   │   request.py
│   │   │   response.py
│   │   │   __init__.py
│   │           
│   ├───services
│   │   │   batch_forecast.py
│   │   │   forecast_engine.py
│   │   │   ml_forecast_engine.py
│   │   │   __init__.py
