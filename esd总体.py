import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.stats as stats

def test_stat(y, iteration):
    std_dev = np.std(y)
    avg_y = np.mean(y)

    if std_dev == 0:
        cal = 0  # Set test statistic to zero if standard deviation is zero
        max_index = 0  # Set max_index to 0 as a placeholder
    else:
        abs_val_minus_avg = abs(y - avg_y)
        max_of_deviations = max(abs_val_minus_avg)
        max_index = np.argmax(abs_val_minus_avg)
        cal = max_of_deviations / std_dev

    print('Test {}'.format(iteration))
    print("Test Statistics Value(R{}) : {}".format(iteration, cal))
    return cal, max_index
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
    '总价值创作用户': [2728323, 2738545, 2765341, 2778084, 2750615, 2731766, 2739661, 2731796, 2724211, 2737845, 2741999, 2757990, 2734544, 2769096, 2764160,
              2719736, 2753054, 2798712, 2772798, 2730842, 2732076, 2743182, 2740714, 2771564, 2776500, 2988748, 2991216, 3040576, 2950494],
    '总价值创作留存用户': [1709860, 1687867, 1693818, 1687648, 1712428, 1667904, 1658032, 1686414, 1693818, 1645692, 1670372, 1685644, 1699218, 1732536, 1715260,
                 1710324, 1670836, 1679474, 1693048, 1702920, 1647390, 1636284, 1710324, 1638752, 1675772, 2050414, 2021095, 2045355, 1987604],
    '总价值创作留存': [63, 62, 61, 61, 62, 61, 61, 62, 62, 60, 61, 61, 62, 63, 62, 63, 61, 60, 61, 62, 60, 60, 62, 59, 60, 69, 68, 67, 67],
}

df = pd.DataFrame(data)

# Apply ESD_Test to each column
for column in df.columns:
    print(f"Outlier detection for {column}:")
    max_i = 0  # Initialize max_i
    result_df = ESD_Test(df[column].values, 0.05, 29)  # 2 is the maximum number of outliers to be detected
    result_df.to_excel(f"{column}_outlier_detection_result.xlsx")  # Save the result to an Excel file
    print("\n" + "=" * 50 + "\n")
