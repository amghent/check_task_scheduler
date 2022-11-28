conda activate base
conda create -y --name check_task_scheduler
conda activate check_task_scheduler
conda install -y -c conda-forge python=3.10 git pyyaml pywin32 croniter
