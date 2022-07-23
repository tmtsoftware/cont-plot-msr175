#!/usr/bin/env python3

import argparse
import os
from pathlib import Path
import sys

import numpy as np
import matplotlib.pyplot as plt

plt.style.use(os.path.join(os.path.dirname(__file__), 'mplstyle'))

def decompose_msr175_data(msr175_data):
    t_ms = msr175_data[:, 0]
    x_g  = msr175_data[:, 1]
    y_g  = msr175_data[:, 2]
    z_g  = msr175_data[:, 3]
    return t_ms, x_g, y_g, z_g

def plot_time_series(msr175_data,
                     output_file,
                     show_total = True,
                     show_max   = True,
                     min_acc    = float('nan'),
                     max_acc    = float('nan')):
    
    t_ms, x_g, y_g, z_g = decompose_msr175_data(msr175_data)
    
    fig = plt.figure(str(output_file))
    ax  = fig.add_subplot()

    ax.plot(t_ms, x_g, label = 'X')
    ax.plot(t_ms, y_g, label = 'Y')
    ax.plot(t_ms, z_g, label = 'Z')
    
    if show_total:
        total_g = np.sqrt(x_g ** 2 + y_g ** 2 + z_g ** 2)
        ax.plot(t_ms, total_g, label = 'Total')

    ax.set_xlabel('Time [ms]')
    ax.set_ylabel('Acceleration [g]')
    ax.set_xlim(( t_ms[0], t_ms[-1] ))
    ax.legend(loc = 'lower right')
    
    if show_total and show_max:
        max_total_g = max(total_g)

        xaxis_max = ax.get_xlim()[1]
        yaxis_max = ax.get_ylim()[1]
        
        ax.text(xaxis_max * 0.98, yaxis_max * 0.95,
                f'Max: {max_total_g:.1f} g',
                horizontalalignment = 'right',
                verticalalignment   = 'top',
                bbox = dict(facecolor = 'white',
                            edgecolor = 'black',
                            boxstyle  = 'round'))

    if not (np.isnan(min_acc) and np.isnan(max_acc)):
        current_ylim = ax.get_ylim()
        new_ylim = (current_ylim[0] if np.isnan(min_acc) else min_acc,
                    current_ylim[1] if np.isnan(max_acc) else max_acc)
        ax.set_ylim(new_ylim)

    fig.savefig(output_file)

def get_sampling_period_in_milliseconds(msr175_data):
    t_ms, x_g, y_g, z_g = decompose_msr175_data(msr175_data)
    return t_ms[1] - t_ms[0]

def read_msr175_csv(csv_file_path):
    msr175_data = np.loadtxt(csv_file_path,
                             dtype     = float,
                             skiprows  = 1,
                             usecols   = (0, 1, 2, 3),
                             delimiter = ',',
                             quotechar = '"')
    return msr175_data

def parse_arguments():
    parser = argparse.ArgumentParser(
        description = 'Tool to plot MSR 175 acceleration data.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    
    parser.add_argument('--plot-format',
                        dest    = 'plot_format',
                        default = 'png',
                        metavar = 'EXT',
                        help    = 'Extension of the image files for plots.')
    parser.add_argument('--hide-total',
                        dest    = 'hide_total',
                        action  = 'store_true',
                        help    = 'Hide total acceleration in the time series plot.')
    parser.add_argument('--hide-max',
                        dest    = 'hide_max',
                        action  = 'store_true',
                        help    = 'Hide maximum accleration in the time series plot.')
    parser.add_argument('--show-plots',
                        dest    = 'show_plots',
                        action  = 'store_true',
                        help    = 'Show the generated plots in GUI')
    parser.add_argument('--min-acc',
                        dest    = 'min_acc',
                        type    = float,
                        default = float('nan'),
                        help    = 'Minimum acceleration in g for the time series plot. Specify "nan" for auto scale.')
    parser.add_argument('--max-acc',
                        dest    = 'max_acc',
                        type    = float,
                        default = float('nan'),
                        help    = 'Maximum acceleration in g for the time series plot. Specify "nan" for auto scale.')
    
    parser.add_argument('csv_file', nargs='+')
    
    args = parser.parse_args()
    return args

def main():
    args = parse_arguments()

    csv_paths = [Path(cf) for cf in args.csv_file]

    for csv_path in csv_paths:
        filename_base = csv_path.stem
        directory     = csv_path.parent

        msr175_data   = read_msr175_csv(csv_path)
        print(f'Loaded data from {csv_path}')

        time_series_plot_path = directory.joinpath(f'{filename_base}.{args.plot_format}')
        
        plot_time_series(msr175_data,
                         time_series_plot_path,
                         show_total = not args.hide_total,
                         show_max   = not args.hide_max,
                         min_acc    = args.min_acc,
                         max_acc    = args.max_acc)
        print(f'Generated time series plot as {time_series_plot_path}')

    if args.show_plots:
        plt.show()

if __name__ == "__main__":
    main()
