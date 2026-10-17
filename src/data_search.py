"""Search imported data without changing selection or generation scope."""


class SearchCancelled(RuntimeError):
    pass


def search_records(dataset, query, cancel=None, progress=None):
    query = query.strip().casefold()
    matches = []
    total = len(dataset.records)
    progress_step = max(200, (total + 99) // 100)
    for index, record in enumerate(dataset.records):
        if index % 200 == 0:
            if cancel is not None and cancel.is_set():
                raise SearchCancelled('搜索已停止，保留原来的显示和选择。')
        if progress and index % progress_step == 0:
            progress(index, total)
        if (not query or query in str(dataset.row_numbers[index])
                or any(query in value.casefold() for value in record.values())):
            matches.append(index)
    if cancel is not None and cancel.is_set():
        raise SearchCancelled('搜索已停止，保留原来的显示和选择。')
    if progress:
        progress(total, total)
    return matches
