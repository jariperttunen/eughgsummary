#!/bin/bash
#SBATCH --job-name=eulandtransition  # create a short name for your job
#SBATCH --account=jarip          # account
#SBATCH --nodes=1                # node count
#SBATCH --ntasks=1              # total number of tasks across all nodes
#SBATCH --cpus-per-task=1        # cpu-cores per task (>1 if multi-threaded tasks)
#SBATCH --mem-per-cpu=4G         # memory per cpu-core
#SBATCH --partition=normal         # sorvi: normal, big or test
#SBATCH --time=7:00:00          # total run time limit (HH:MM:SS)
#SBATCH --output=output_%j.txt
#SBATCH --error=errors_%j.txt
#SBATCH --mail-type=begin        # send email when job begins
#SBATCH --mail-type=end          # send email when job ends
#SBATCH --mail-type=FAIL         # sen email on fail
#SBATCH --mail-user=jari.perttunen@luke.fi

#Activate virtual environment
source /home/jarip/venv/euco2/bin/activate
#Data in CRF_vertailu and collect data for EU plus GBR, ISL and NOR
srun  python eulandusechangeCO2.py -d /data/projects/khk/CRF_vertailu_2023/ -s 1990 -e 2021 --list &
wait
