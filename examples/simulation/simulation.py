"""
Copyright (C) 2014, Zoomer Analytics LLC.
All rights reserved.

Version: 0.1.0
License: BSD 3-clause (see LICENSE.txt for details)
"""
from __future__ import division
import numpy as np
from xlwings import Workbook, Range, Chart

wb = Workbook()

def main():
    # User Inputs
    num_simulations = int(Range('E3').value)
    time = Range('E4').value
    num_timesteps = int(Range('E5').value)
    dt = time/num_timesteps  # Length of time period
    vol = Range('E7').value
    mu = np.log(1 + Range('E6').value)  # Drift
    starting_price = Range('E8').value
    perc_selection = [5, 50, 95]  # percentiles (hardcoded for now)
    # Animation
    if wb.com_workbook.ActiveSheet.OLEObjects("ComboBox1").Object.Value == 'Yes':
        animate = True
    else:
        animate = False

    # Excel: clear output, write out initial values of percentiles/sample path and set chart source
    # and x-axis values
    Range('O2').table.clear_contents()
    Range('P2').value = [starting_price, starting_price, starting_price, starting_price]
    Chart('Chart 5').set_source_data(Range((1, 15),(num_timesteps + 2, 19)))
    Range('O2').value = np.round(np.linspace(0, time, num_timesteps + 1).reshape(-1,1), 2)

    # Preallocation
    price = np.zeros((num_timesteps + 1, num_simulations))
    percentiles = np.zeros((num_timesteps + 1, 3))

    # Set initial values
    price[0,:] = starting_price
    percentiles[0,:] = starting_price

    # Simulation at each time step
    for t in range(1, num_timesteps + 1):
        rand_nums = np.random.randn(num_simulations)
        price[t,:] = price[t-1,:] * np.exp((mu - 0.5 * vol**2) * dt + vol * rand_nums * np.sqrt(dt))
        percentiles[t,:] = np.percentile(price[t,:], perc_selection)
        if animate:
            Range((t+2,16)).value = percentiles[t,:]
            Range((t+2,19)).value = price[t,0]  # Sample path
            wb.com_app.Application.ScreenUpdating = True

    if not animate:
        Range('P2').value = percentiles
        Range('S2').value = price[:,:1]  # Sample path

if __name__ == '__main__':
    main()



