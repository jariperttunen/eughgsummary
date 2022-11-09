#!/bin/bash
#SBATCH --job-name=eughgsummary   # create a short name for your job
#SBATCH --account=jarip          # account
#SBATCH --nodes=1                # node count
#SBATCH --ntasks=3              # total number of tasks across all nodes
#SBATCH --cpus-per-task=1        # cpu-cores per task (>1 if multi-threaded tasks)
#SBATCH --mem-per-cpu=4G         # memory per cpu-core
#SBATCH --partition=normal         # sorvi: normal, big or test
#SBATCH --time=04:00:00          # total run time limit (HH:MM:SS)
#SBATCH --output=output_%j.txt
#SBATCH --error=errors_%j.txt
#SBATCH --mail-type=begin        # send email when job begins
#SBATCH --mail-type=end          # send email when job ends
#SBATCH --mail-user=jari.perttunen@luke.fi

#Activate virtual environment
source /home/jarip/venv/euco2/bin/activate
#Data in CRF_vertailu and collect data for EU plus GBR, ISL and NOR
srun -N 1 -n 1 python eurestoration.py -d /data/d4/projects/khk/CRF_vertailu_2022/ -s 1990 -e 2019 --eu &
wait
