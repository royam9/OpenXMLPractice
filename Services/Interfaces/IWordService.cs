﻿namespace Services.Interfaces;

/// <summary>
/// Word相關服務
/// </summary>
public interface IWordService
{
    Task AddExcelChartToExistingWordDocument(string filePath);
}
