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
    '低价值创作用户': [1229000, 1241559, 1243685, 1239906, 1226632, 1231238, 1257593, 1235050, 1225087, 1227465,
             1257593, 1234000, 1246340, 1248808, 1245106, 1231532, 1236468, 1262382, 1240170, 1230298,
             1232766, 1257446, 1245106, 1254978, 1254978, 1234000, 1231532, 1232766, 1231532],
    '低价值创作留存用户': [653956, 656643, 673700, 658892, 657758, 642850, 642850, 653956, 661360, 615702,
               653956, 658956, 661424, 678700, 663892, 662658, 647850, 647850, 658956, 666360,
               620702, 620702, 662658, 623170, 645382, 802100, 795930, 771250, 783590],
    '低价值创作留存': [53, 53, 54, 53, 54, 52, 51, 53, 54, 50,
             52, 53, 53, 54, 53, 54, 52, 51, 53, 54,
             50, 49, 53, 50, 51, 65, 65, 63, 64]
}

df = pd.DataFrame(data)

# Apply ESD_Test to each column
for column in df.columns:
    print(f"Outlier detection for {column}:")
    max_i = 0  # Initialize max_i
    result_df = ESD_Test(df[column].values, 0.05, 29)  # 2 is the maximum number of outliers to be detected
    result_df.to_excel(f"{column}_outlier_detection_result.xlsx")  # Save the result to an Excel file
    print("\n" + "=" * 50 + "\n")
