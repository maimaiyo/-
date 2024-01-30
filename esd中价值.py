import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.stats as stats

def test_stat(y, iteration):
    std_dev = np.std(y)
    avg_y = np.mean(y)
    abs_val_minus_avg = abs(y - avg_y)
    max_of_deviations = max(abs_val_minus_avg)
    max_ind = np.argmax(abs_val_minus_avg)
    cal = max_of_deviations / std_dev
    print('Test {}'.format(iteration))
    print("Test Statistics Value(R{}) : {}".format(iteration, cal))
    return cal, max_ind

def calculate_critical_value(size, alpha, iteration):
    t_dist = stats.t.ppf(1 - alpha / (2 * size), size - 2)
    numerator = (size - 1) * np.sqrt(np.square(t_dist))
    denominator = np.sqrt(size) * np.sqrt(size - 2 + np.square(t_dist))
    critical_value = numerator / denominator
    print("Critical Value(λ{}): {}".format(iteration, critical_value))
    return critical_value

def check_values(R, C, inp, max_index, iteration):
    if R > C:
        print('{} is an outlier. R{} > λ{}: {:.4f} > {:.4f} \n'.format(inp[max_index], iteration, iteration, R, C))
    else:
        print('{} is not an outlier. R{} > λ{}: {:.4f} > {:.4f} \n'.format(inp[max_index], iteration, iteration, R, C))

def ESD_Test(input_series, alpha, max_outliers):
    stats = []
    critical_vals = []
    for iterations in range(1, max_outliers + 1):
        stat, max_index = test_stat(input_series, iterations)
        critical = calculate_critical_value(len(input_series), alpha, iterations)
        check_values(stat, critical, input_series, max_index, iterations)
        input_series = np.delete(input_series, max_index)
        critical_vals.append(critical)
        stats.append(stat)

    print('H0:  there are no outliers in the data')
    print('Ha:  there are up to 10 outliers in the data')
    print('')
    print('Significance level:  α = {}'.format(alpha))
    print('Critical region:  Reject H0 if Ri > critical value')
    print('Ri: Test statistic')
    print('λi: Critical Value')
    print(' ')
    df = pd.DataFrame({'i': range(1, max_outliers + 1), 'Ri': stats, 'λi': critical_vals})

    def highlight_max(x):
        if x.i == max_i:
            return ['background-color: yellow'] * 3
        else:
            return ['background-color: white'] * 3

    df.index = df.index + 1
    print('Number of outliers {}'.format(max_i))

    return df.style.apply(highlight_max, axis=1)

# Your data
data = {
    '中价值创作用户': [867615, 877508, 883878, 886746, 878721, 879946, 863954, 876164, 881010, 876224, 863824, 869970, 869970, 881076, 887246,
              868736, 878608, 884778, 887246, 879842, 881076, 865034, 877374, 882310, 877374, 866268, 886012, 882310, 867502],
    '中价值创作留存用户': [603026, 586984, 584516, 582048, 603026, 591920, 580814, 598090, 583282, 598090, 583282, 592320, 584916, 598490, 598490,
                 603426, 587384, 584916, 582448, 603426, 592320, 581214, 598490, 583682, 598490, 599724, 587384, 598490, 582448],
    '中价值创作留存': [70, 67, 66, 66, 69, 67, 67, 68, 66, 68, 68, 68, 67, 68, 67, 69, 67, 66, 66, 69, 67, 67, 68, 66, 68, 69, 66, 68, 67],
}

df = pd.DataFrame(data)

# Apply ESD_Test to each column
for column in df.columns:
    print(f"Outlier detection for {column}:")
    max_i = 0  # Initialize max_i
    result_df = ESD_Test(df[column].values, 0.05, 29)  # 2 is the maximum number of outliers to be detected
    result_df.to_excel(f"{column}_outlier_detection_result.xlsx")  # Save the result to an Excel file
    print("\n" + "=" * 50 + "\n")
