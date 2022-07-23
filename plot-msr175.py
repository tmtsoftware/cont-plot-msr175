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

def plot_time_series(ax,
                     msr175_data,
                     show_total = True,
                     show_max   = True,
                     acc_min_g  = float('nan'),
                     acc_max_g  = float('nan'),
                     t_min_ms   = 0.0,
                     t_max_ms   = float('nan')):
    
    t_ms, x_g, y_g, z_g = decompose_msr175_data(msr175_data)
    
    ax.plot(t_ms, x_g, label = 'X')
    ax.plot(t_ms, y_g, label = 'Y')
    ax.plot(t_ms, z_g, label = 'Z')
    
    if show_total:
        total_g = np.sqrt(x_g ** 2 + y_g ** 2 + z_g ** 2)
        ax.plot(t_ms, total_g, label = 'Total')

    ax.set_xlabel('Time [ms]')
    ax.set_ylabel('Acceleration [g]')
    ax.legend(loc = 'lower right')
    
    # Set X axis range.
    ax.set_xlim(( t_min_ms,
                  t_ms[-1] if np.isnan(t_max_ms) else t_max_ms ))

    # Set Y axis range.
    if not (np.isnan(acc_min_g) and np.isnan(acc_max_g)):
        current_ylim = ax.get_ylim()
        new_ylim = (current_ylim[0] if np.isnan(acc_min_g) else acc_min_g,
                    current_ylim[1] if np.isnan(acc_max_g) else acc_max_g)
        ax.set_ylim(new_ylim)

    # Show max acceleration as text in the plot.
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

def get_sampling_period_in_milliseconds(msr175_data):
    t_ms, x_g, y_g, z_g = decompose_msr175_data(msr175_data)
    return t_ms[1] - t_ms[0]

def plot_power_spectrum(ax,
                        msr175_data,
                        ps_min_g2 = float('nan'),
                        ps_max_g2 = float('nan')):
    t_ms, x_g, y_g, z_g = decompose_msr175_data(msr175_data)

    dt_ms = get_sampling_period_in_milliseconds(msr175_data)

    x_ps = np.abs(np.fft.fft(x_g))**2
    y_ps = np.abs(np.fft.fft(y_g))**2
    z_ps = np.abs(np.fft.fft(z_g))**2

    assert len(x_g) == len(y_g)
    assert len(y_g) == len(z_g)
    n = len(x_g)

    freq_Hz = np.fft.fftfreq(n, dt_ms / 1000.0)

    x_ps = x_ps[0:int(n/2)]
    y_ps = y_ps[0:int(n/2)]
    z_ps = z_ps[0:int(n/2)]
    freq_Hz = freq_Hz[0:int(n/2)]
    
    ax.plot(freq_Hz, x_ps, label = 'X')
    ax.plot(freq_Hz, y_ps, label = 'Y')
    ax.plot(freq_Hz, z_ps, label = 'Z')
    
    ax.set_xlabel('Frequency [Hz]')
    ax.set_ylabel(r'Power Spectrum [${\rm g}^2$]')
    ax.legend(loc = 'upper right')
    
    # Set X axis range.
    ax.set_xlim((0, freq_Hz[-1]))

    # Set Y axis.
    ax.set_yscale('log')
    if not (np.isnan(ps_min_g2) and np.isnan(ps_max_g2)):
        current_ylim = ax.get_ylim()
        new_ylim = (current_ylim[0] if np.isnan(ps_min_g2) else ps_min_g2,
                    current_ylim[1] if np.isnan(ps_max_g2) else ps_max_g2)
        ax.set_ylim(new_ylim)
        
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
    parser.add_argument('--dpi',
                        dest    = 'dpi',
                        type    = float,
                        default = 96.0,
                        help    = 'DPI value for plot rendering.')
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
                        help    = 'Show the generated plots in GUI.')
    parser.add_argument('--min-acc',
                        dest    = 'acc_min_g',
                        type    = float,
                        default = float('nan'),
                        help    = 'Minimum acceleration in g for the time series plot. Specify "nan" for auto scale.')
    parser.add_argument('--max-acc',
                        dest    = 'acc_max_g',
                        type    = float,
                        default = float('nan'),
                        help    = 'Maximum acceleration in g for the time series plot. Specify "nan" for auto scale.')
    parser.add_argument('--min-time',
                        dest     = 't_min_ms',
                        type     = float,
                        default  = 0.0,
                        help     = 'Minimum time in the time series plot in milliseconds.')
    parser.add_argument('--max-time',
                        dest     = 't_max_ms',
                        type     = float,
                        default  = float('nan'),
                        help     = 'Maximum time in the time series plot in milliseconds. Specify "nan" for auto scale.')
    parser.add_argument('--plot-power-spectrum',
                        dest    = 'plot_power_spectrum',
                        action  = 'store_true',
                        help    = 'Plow power spectrum.')
    parser.add_argument('--min-ps',
                        dest    = 'ps_min_g2',
                        type    = float,
                        default = float('nan'),
                        help    = 'Minimum power spectrum in g^2 for the plot. Specify "nan" for auto scale.')
    parser.add_argument('--max-ps',
                        dest    = 'ps_max_g2',
                        type    = float,
                        default = float('nan'),
                        help    = 'Maximum power spectrum in g^2 for the plot. Specify "nan" for auto scale.')
    
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

        plot_path = directory.joinpath(f'{filename_base}.{args.plot_format}')
        fig = plt.figure(str(plot_path))
        
        if args.plot_power_spectrum:
            ax_time_series  = fig.add_subplot(2, 1, 1)
        else:
            ax_time_series  = fig.add_subplot(1, 1, 1)
        
        plot_time_series(ax_time_series,
                         msr175_data,
                         show_total = not args.hide_total,
                         show_max   = not args.hide_max,
                         acc_min_g  = args.acc_min_g,
                         acc_max_g  = args.acc_max_g,
                         t_min_ms   = args.t_min_ms,
                         t_max_ms   = args.t_max_ms)

        if args.plot_power_spectrum:
            ax_power_spectrum = fig.add_subplot(2, 1, 2)
            plot_power_spectrum(ax_power_spectrum,
                                msr175_data,
                                ps_min_g2 = args.ps_min_g2,
                                ps_max_g2 = args.ps_max_g2)

        fig.savefig(plot_path, dpi = args.dpi)
        print(f'Generated the plot as {plot_path}')

    if args.show_plots:
        plt.show()

if __name__ == "__main__":
    main()
