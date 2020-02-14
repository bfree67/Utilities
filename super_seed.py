# -*- coding: utf-8 -*-
"""
Created on Fri Feb 14 13:29:47 2020
A super randomizer

@author: Brian
"""

import numpy as np
import time

def super_seed(n=1000):
    ### make a random seed (assumes you imported numpy as np and time)
    seed = int(time.time())
    np.random.seed(seed) # first seed from the system clock
    
    #n is a length of a random list of integers and select one from random
    x_seed = np.random.uniform(0,1000,n).astype(int) # make a list of 100 different random variables
    x_index = int(np.random.uniform(n))
    
    np.random.seed(x_seed[x_index]) # second seed
    
super_seed()

